<!-- #include file="../conn.asp" -->
<!-- #include file="init.asp" -->
<%

 ndate = request("stdate1")
 yy=right(ndate,4)
 mm=mid(ndate,4,2)
 dd = left(ndate,2)
xxndate = yy&"/"&right("00"&mm,2)&"/"&right("00"&dd,2)


server.scripttimeout = 1800
xdate =dateserial(yy,mm,dd)
mdate = xdate

ydate =dateserial(yy,mm,1)
ttlshamt = 0
ttllnamt = 0
dim  acnt(6)
dim shcnt(9)
dim lncnt(9)
for i = 1 to 9
    lncnt(i) = 0
next
for i = 1 to 9
    shcnt(i) = 0
next
for i = 1 to 6
    acnt(i) = 0
next 
cnt = 0



set rs = conn.execute( "select a.memno,a.code,a.amount ,b.membday "&_
                       " from share a,memmaster b where a.memno=b.memno and "&_
                       "a.ldate<='"&xxndate&"' order by a.memno , a.code ")                                  
                     
                     
xmemno = rs("memno")
membday = rs("membday")
xamt = 0
bbb = 0
do while not rs.eof
             
            if xmemno <> rs("memno") then
 
               if xamt > 0 then  
                  bbb = bbb +1
 
 
                 xmembday =  membday

                 ttlshamt = ttlshamt + xamt
             if xmembday<>"" then
                 membday = right("0"&day(xmembday),2)&"/"&right("0"&month(xmembday),2)&"/"&year(xmembday)
                  age   = year(mdate) -  year(xmembday)
                 bdate = dateserial( yy , month(xmembday), day(xmembday) )
                 nday = (mdate - bdate )
             
            
                 
                  if nday  < 0  then
                     age = age - 1 
                  end if
                  membday = rs("membday")                     
            else
                  
                  age = 0
                  membday=""  
            end if
 

 
          
               
                 if age>=18 and age <=30 then
                     acnt(1) = acnt(1) + 1
                  end if    
              if age>=31 and age <=40 then
                 acnt(2) = acnt(2) + 1    
              end if       
              if age>=41 and age <=50 then
                 acnt(3) = acnt(3) + 1
              end if    
              if age>=51 and age <=60 then
                acnt(4) = acnt(4) + 1
              end if  
              if age>=61 and age <=70 then
                  acnt(5) = acnt(5) + 1
              end if    
             if age>=71  then
                acnt(6) = acnt(6) + 1
              end if  
                
              if xamt >=5 and xamt <=50000 then
                 shcnt(1) = shcnt(1) + 1
              end if
              if xamt >=50001 and xamt <=100000 then
                 shcnt(2)  = shcnt(2) + 1
             end if
              if xamt >=100001 and xamt <=150000 then
                 shcnt(3)  = shcnt(3) + 1
              end if 
              if xamt >=150001 and xamt <=200000 then
                 shcnt(4)  = shcnt(4) + 1
              end if
              if xamt >=200001 and xamt <=250000 then
                 shcnt(5) = shcnt(5) + 1
              end if
              if xamt >=250001 and xamt <=300000 then
                 shcnt(6)  = shcnt(6) + 1
             end if
              if xamt >=300001 and xamt <=350000 then
                 shcnt(7)  = shcnt(7) + 1
              end if 
              if xamt >=350001 and xamt <=400000 then
                 shcnt(8)  = shcnt(8) + 1
              end if
             if xamt >=400001 and xamt <=660000  then
                 shcnt(9)  = shcnt(9) + 1
              end if
              end if 
               xmemno = rs("memno") 
               membday = rs("membday") 
               xamt = 0 
            end if
           
               select case rs("code")
                      case "0A","A1","A2","A3","C0","C1","C3","A0","A4","A7"
                            
                    xamt = xamt + rs("amount")
                 
                  case "B0","B1","G0","G1","G3","H0","H1","H3","MF"
                xamt = xamt - rs("amount")
         end select
      
              
    rs.movenext
loop
rs.close


             if xamt > 0 then  
 
                  ttlshamt = ttlshamt + xamt
                xmembday =  membday  

             if xmembday<>"" then
                 membday = right("0"&day(xmembday),2)&"/"&right("0"&month(xmembday),2)&"/"&year(xmembday)
                  age   = year(mdate) -  year(xmembday)
                 bdate = dateserial( yy , month(xmembday), day(xmembday) )
                 nday = (mdate - bdate )
             
                  if nday  < 0  then
                     age = age - 1 
                  end if
                                    
            else
                  
                  age = 0
                  membday=""  
            end if
                      
                 if age>=18 and age <=30 then
                     acnt(1) = acnt(1) + 1
                  end if    
              if age>=31 and age <=40 then
                 acnt(2) = acnt(2) + 1    
              end if       
              if age>=41 and age <=50 then
                 acnt(3) = acnt(3) + 1
              end if    
              if age>=51 and age <=60 then
                acnt(4) = acnt(4) + 1
              end if  
              if age>=61 and age <=70 then
                  acnt(5) = acnt(5) + 1
              end if    
             if age>=71  then
                acnt(6) = acnt(6) + 1
              end if  
              if xamt >=5 and xamt <=50000 then
                 shcnt(1) = shcnt(1) + 1
              end if
              if xamt >=50001 and xamt <=100000 then
                 shcnt(2)  = shcnt(2) + 1
             end if
              if xamt >=100001 and xamt <=150000 then
                 shcnt(3)  = shcnt(3) + 1
              end if 
              if xamt >=150001 and xamt <=200000 then
                 shcnt(4)  = shcnt(4) + 1
              end if
              if xamt >=200001 and xamt <=250000 then
                 shcnt(5) = shcnt(5) + 1
              end if
              if xamt >=250001 and xamt <=300000 then
                 shcnt(6)  = shcnt(6) + 1
             end if
              if xamt >=300001 and xamt <=350000 then
                 shcnt(7)  = shcnt(7) + 1
              end if 
              if xamt >=350001 and xamt <=400000 then
                 shcnt(8)  = shcnt(8) + 1
              end if
             if xamt >=400001 and xamt <=660000 then
                 shcnt(9)  = shcnt(9) + 1
              end if
           end if   
           set ms = server.createobject("ADODB.Recordset")
 
                lnamt = 0  
              mssql = "select a.memno,a.code,a.amount ,a.ldate ,a.lnnum from loan a ,memmaster b where a.memno=b.memno and b.mstatus <>'V'  and a.ldate <= '"&xxndate&"'  order by a.memno,a.ldate,right(a.code ,1),left(a.code,1) "


                ms.open mssql, conn, 1, 1              
                xmemno = ms("memno")
                 if not ms.eof then
                do while not ms.eof 
                   if xmemno<> ms("memno")   then

                if lnamt > 0 then

                       ttllnamt = ttllnamt + lnamt
                          if lnamt = 0 then
                             lncnt(1)= lncnt(1) + 1
                          end if
                          if lnamt>=1     and lnamt <=50000 then 
                             lncnt(2) = lncnt(2) + 1
                          end if
                          if lnamt>=50001 and lnamt <=100000 then
                             lncnt(3) = lncnt(3) + 1
                          end if
                                 if lnamt>=100001 and lnamt <=150000 then
                                    lncnt(4) = lncnt(4) + 1
                                 end if
                                 if lnamt>=150001 and lnamt <=200000 then
                                    lncnt(5) = lncnt(5) + 1
                                  end if
                          if lnamt >=200001 and lnamt <=250000 then
                             lncnt(6) = lncnt(6) + 1
                          end if
                          if lnamt >=250001 and lnamt <=300000 then
                             lncnt(7) = lncnt(7) + 1
                          end if  
                          if lnamt >=300001 and lnamt <= 320000 then
                             lncnt(8) = lncnt(8) + 1
                          end if
                
                      end if   
 

                      xmemno= ms("memno")
                      lnamt = 0
                   end if
                  
                      select case ms("code")
                             case "0D"
 
                                      lnamt = lnamt + ms("amount")
                    case "D9"
                             set ms1 =  conn.execute("select appamt from loanrec where lnnum= '"&ms("lnnum")&"' ")
                                 if not ms1.eof then
                                        lnamt= lnamt + ms1(0)
                                  end if
                               ms1.close                          
                             case "D8","E0","E1","E2","E3","E6","E7","EC"
                                 
                                      lnamt = lnamt - ms("amount")
                              case "DF"
                                     if ms("memno")="4480" and  ymd(ms("ldate")) >="2016/07/30" and ms("lnnum")= "2013080003" then    
                                        lnamt = lnamt + ms("amount")
   
                                         end if        


                     end select
                   
                    ms.movenext
                    loop  
                  if lnamt > 0  then

                       ttllnamt = ttllnamt + lnamt
                          if lnamt = 0 then
                             lncnt(1)= lncnt(1) + 1
                          end if
                          if lnamt>=1     and lnamt <=50000 then 
                             lncnt(2) = lncnt(2) + 1
                          end if
                          if lnamt>=50001 and lnamt <=100000 then
                             lncnt(3) = lncnt(3) + 1
                          end if
                                 if lnamt>=100001 and lnamt <=150000 then
                                    lncnt(4) = lncnt(4) + 1
                                 end if
                                 if lnamt>=150001 and lnamt <=200000 then
                                    lncnt(5) = lncnt(5) + 1
                                  end if
                          if lnamt >=200001 and lnamt <=250000 then
                             lncnt(6) = lncnt(6) + 1
                          end if
                          if lnamt >=250001 and lnamt <=300000 then
                             lncnt(7) = lncnt(7) + 1
                          end if  
                          if lnamt >=300001 and lnamt <= 320000 then
                             lncnt(8) = lncnt(8) + 1
                          end if
                
                      end if   
                 end if              
                 ms.close 
          
  
                   
 
 
 ttlacnt = 0 
 for i = 1 to 6
     ttlacnt = ttlacnt + acnt(i)
 next 
ttlshcnt = 0 
 for i = 1 to 9
     ttlshcnt = ttlshcnt + shcnt(i)
 next 
ttllncnt = 0 
 for i = 2 to 8
      
     ttllncnt = ttllncnt + lncnt(i)
 next 
mndate=right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
xdate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())


         
 
if request.form("output")="word" then
	Response.ContentType = "application/msword"
elseif request.form("output")="excel" then
	Response.ContentType = "application/vnd.ms-excel"

end if
%>
<html>
<head>
<title>社員統計資料分部報告</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body rightMargin="10" topmargin="10" marginheight="0" marginwidth="0">
<center>
<table border="0" cellpadding="0" cellspacing="0">
     
	<tr height="30" valign="top" align="center"><td colspan="15"><font size="3"  face="標楷體" >水務署員工儲蓄互助社<br><font size="2"  face="標楷體" >社員統計資料分部報告<br>截至：<%=ndate%></font></font></td></tr>
        <tr height="30" ><td colspan=9></td></tr>

</table>
<table border="1" cellspacing="1" cellpadding="2" align="center"  >
	<tr >
             <td width="200" align="left"><font size="3"  face="標楷體" >（一）社員年齡</font></td>
             <td width="300" align="center"><font size"3"  face="標楷體" >年齡</font></td>
             <td with="100"  align=""centert"><font size="3"  face="標楷體" >人數</font></td>

	</tr>
        <tr>
             <td>&nbsp;</td>
             <td width="300" align="center">18-30</td>
             <td with="100"  align="right"><%=formatNumber(acnt(1),0)%></td> 
        </tr> 
        <tr>
             <td>&nbsp;</td>
             <td width="300" align="center">31-40</td>
             <td with="100"  align="right"><%=formatNumber(acnt(2),0)%></td> 
        </tr> 
        <tr>
             <td>&nbsp;</td>
             <td width="300" align="center">41-50</td>
             <td with="100"  align="right"><%=formatNumber(acnt(3),0)%></td> 
        </tr> 
        <tr>
            <td>&nbsp;</td>
             <td width="300" align="center">51-60</td>
             <td with="100"  align="right"><%=formatNumber(acnt(4),0)%></td> 
        </tr>
        <tr>
             <td>&nbsp;</td>
             <td width="300" align="center">61-70</td>
             <td with="100"  align="right"><%=formatNumber(acnt(5),0)%></td> 
        </tr> 
        <tr>
            <td>&nbsp;</td>
             <td width="300" align="center">71＞</td>
             <td with="100"  align="right"><%=formatNumber(acnt(6),0)%></td> 
        </tr>  
        <tr>
             <td>&nbsp;</td>
             <td width="300" align="right"><font size="3"  face="標楷體" >總人數</font></td>
             <td with="100"  align="right"><%= formatNumber(ttlacnt,0)%></td> 
        </tr> 
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>  
	<tr>
             <td width="200" align="left"><font size="3"  face="標楷體" >（二）社員平均儲蓄</font></td>
             <td width="300" align="center"><font size"3"  face="標楷體" >股金</font></td>
             <td with="100"  align="center"><font size="3"  face="標楷體" >人數</font></td>

	</tr>
        <tr>
             <td>&nbsp;</td>
             <td width="300" align="center">$5-$50,000</td>
             <td with="100"  align="right"><%=formatNumber(shcnt(1),0)%></td> 
        </tr> 
        <tr>
             <td>&nbsp;</td>
             <td width="300" align="center">$50,001-$100,000</td>
             <td with="100"  align="right"><%=formatNumber(shcnt(2),0)%></td> 
        </tr> 
        <tr>
            <td>&nbsp;</td>
             <td width="300" align="center">$100,001-$150,000</td>
             <td with="100"  align="right"><%=formatNumber(shcnt(3),0)%></td> 
        </tr> 
        <tr>
             <td>&nbsp;</td>
             <td width="300" align="center">$150,001-$200,000</td>
             <td with="100"  align="right"><%=formatNumber(shcnt(4),0)%></td> 
        </tr>
        <tr>
            <td>&nbsp;</td>
             <td width="300" align="center">$200,001-$250,000</td>
             <td with="100"  align="right"><%=formatNumber(shcnt(5),0)%></td> 
        </tr> 
        <tr>
             <td>&nbsp;</td>
             <td width="300" align="center">$250,001-$300,000</td>
             <td with="100"  align="right"><%=formatNumber(shcnt(6),0)%></td> 
        </tr>
        <tr>
             <td>&nbsp;</td>
             <td width="300" align="center">$300,001-$350,000</td>
             <td with="100"  align="right"><%=formatNumber(shcnt(7),0)%></td> 
        </tr>
        <tr>
             <td>&nbsp;</td>
             <td width="300" align="center">$350,001-$400,000</td>
             <td with="100"  align="right"><%=formatNumber(shcnt(8),0)%></td> 
        </tr> 
        <tr>
             <td>&nbsp;</td>
             <td width="300" align="center">$400,001-$660,000</td>
             <td with="100"  align="right"><%=formatNumber(shcnt(9),0)%></td> 
        </tr>    
        <tr>
             <td>&nbsp;</td>
             <td>
           <table border="0" cellspacing="0" cellpadding="2" align="center"  >
                          <td width="80" align="left"><font size="3"  face="標楷體" >股金總額：</font></td> 
                          <td width="120" align="left"><%=formatNumber(ttlshamt,2)%></td>
                          <td width="100" align="right"><font size="3"  face="標楷體" >總人數</font></td>
             </table> 
             </td>
             <td with="100"  align="right"><%=formatNumber( ttlshcnt,0)%></td> 
        </tr> 

  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>  
	<tr>
             <td width="200" align="left"><font size="3"  face="標楷體" >（二）社員平均貸款</font></td>
             <td width="300" align="center"><font size"3"  face="標楷體" >貸款</font></td>
             <td with="100"  align="right"><font size="3"  face="標楷體" >人數</font></td>

	</tr>
        <tr>
             <td>&nbsp;</td>
             <td width="300" align="center">NIL</td>
             <td with="100"  align="right"><%=formatNumber(lncnt(1),0)%></td> 
        </tr> 
        <tr>
             <td>&nbsp;</td>
             <td width="300" align="center">$1-$50,000</td>
             <td with="100"  align="right"><%=formatNumber(lncnt(2),0)%></td> 
        </tr> 
        <tr>
             <td>&nbsp;</td>
             <td width="300" align="center">$50,001-$100,000</td>
             <td with="100"  align="right"><%=formatNumber(lncnt(3),0)%></td> 
        </tr> 
        <tr>
            <td>&nbsp;</td>
             <td width="300" align="center">$100,001-$150,000</td>
             <td with="100"  align="right"><%=formatNumber(lncnt(4),0)%></td> 
        </tr> 
        <tr>
             <td>&nbsp;</td>
             <td width="300" align="center">$150,001-$200,000</td>
             <td with="100"  align="right"><%=formatNumber(lncnt(5),0)%></td> 
        </tr>
        <tr>
            <td>&nbsp;</td>
             <td width="300" align="center">$200,001-$250,000</td>
             <td with="100"  align="right"><%=formatNumber(lncnt(6),0)%></td> 
        </tr> 
        <tr>
             <td>&nbsp;</td>
             <td width="300" align="center">$250,001-$300,000</td>
             <td with="100"  align="right"><%=formatNumber(lncnt(7),0)%></td> 
        </tr>
       <tr>
             <td>&nbsp;</td>
             <td width="300" align="center">$300,001-$320,000</td>
 

             <td with="100"  align="right"><%=formatNumber(lncnt(8),0)%></td> 
        </tr>  
 
        <tr>
             <td>&nbsp;</td>
             <td width="300" align="right">
            <table border="0" cellspacing="0" cellpadding="2" align="center"  >
                          <td width="90" align="left"><font size="3"  face="標楷體" >貸款總額：</font></td> 
                          <td width="110" align="left"><%=formatNumber(ttllnamt,2)%></td>
                          <td width="100" align="right"><font size="3"  face="標楷體" >總人數</font></td>
             </table> 

             </td>
             <td with="100"  align="right"><%=formatNumber(ttllncnt,0)%></td> 
        </tr> 
</table>
</center>
</body>
</html>

