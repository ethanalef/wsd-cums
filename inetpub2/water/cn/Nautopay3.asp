<!-- #include file="../conn.asp" -->

<!-- #include file="../addUserLog.asp" -->
<%

server.scripttimeout = 1800

y=year(date())
m=month(date())

if request("submit")<>"" then
y=year(date())
m=month(date())
D=DAY(date()) 	
SD = 1
	inday = request("mDay")
	mdate = y&"."&m&"."&request("mDay")       
        xDate  = y&"."&right("0"&m,2)
        yydate = y&"."&right("0"&m,2)&".01" 
        Pdate = y&"."&m&"."&d
        STDATE = y&"."&m&"."&SD
        if (y/4)=int(y/4) and (y/100)=int(y/100) then
           MDD = cint(mid("312931303130313130313031",(m-1)*2+1,2))
        else
           mDD = cint(mid("312831303130313130313031",(m-1)*2+1,2)  )
        end if
    
        preyr  = y
        premon = m - 1
        if premon=0 then
           preyr = y - 1 
           premon = 12
        end if 
        ydate = preyr&"."&right("0"&premon,2)
       
        conn.begintrans   
        conn.execute("delete autopay")
        conn.committrans
        conn.begintrans    

        set rs = server.createobject("ADODB.Recordset")
        set Ls = server.createobject("ADODB.Recordset")
        
	sql = "select a.memno,a.mstatus,a.monthsave,a.monthssave,a.tpayamt,a.bnklmt,b.lnnum,b.lndate,b.monthrepay,b.bal,b.delyflag,b.chkmon,b.months,b.chequeamt,b.lnflag  from memmaster a,loanrec b where ( a.mstatus='A' or a.mstatus='T'  or a.mstatus='M' or a.mstatus='0' or  a.mstatus='1' or a.mstatus='2' OR a.mstatus='3' OR a.mstatus='8' )   and   a.wdate is null and a.memno=b.memno and b.repaystat='N' union select memno,mstatus,monthsave,monthssave,tpayamt,bnklmt,null,null,0,0,null,0,0,0,null from memmaster where ( mstatus='A' or mstatus='T'  or mstatus='M' or mstatus='0' or  mstatus='1' or mstatus='2' OR mstatus='3' OR mstatus='8' )   and   wdate is null and memno not in (select memno from loanrec where repaystat='N')     "
        rs.open sql, conn, 1, 1       
        do while not rs.eof
          

           pamt1 = 0 
           pamt2 = 0
           pint1 = 0 
           pint2 = 0
           pint3 = 0
           ttlpamt = 0 
           ttlpint = 0
           samt = 0 
           ttlsamt = 0
           sumttl = 0
           lnflag=rs("lnflag")
           chequeamt = rs("chequeamt")
           memno = rs("memno") 
           mstatus = rs("mstatus")
           bal = rs("bal")
           monthrepay = rs("monthrepay")
           lnnum = rs("lnnum")
           delyflag = rs("delyflag")

 
          if lnnum<>"" then
     
           
              if rs("bal") > rs("monthrepay")  then 
                 pamt1 = rs("monthrepay")
              else
                 pamt1 = rs("bal")
              end if
                
              if rs("delyflag") ="Y" then
                 mstatus = "1"
              end if             
               
              xyy = year(rs("lndate"))
              xmn = month(rs("lndate"))
              xdd = day(rs("lndate"))
              yy  = xyy 
              mn  = xmn
              mdd = mid("312831303130313130313031",(mn-1)*2+1,2)
              if yy/4 = int(yy/4) and yy/100=int(yy/100) then
                 mdd = 29
              end if
 
              chkdate = xyy&"."&right("0"&xmn,2)
              lndate = rs("lndate") 
              sqlstr = "select * from loan where  memno='"&memno&"' and pflag=1 and left(code,1) ='M' "
              Ls.open sqlstr, conn, 2, 2
              do while not ls.eof
                
                 select case ls("code")
                        case "ME"
                             pamt2 = pamt2 +  ls("bal")
                        case "MF"
                             pint2 = pint2 + ls("bal")
                        case "IF"
                             pint2 = pint2 + ls("bal")
                 end select
               ls.movenext
               loop
               ls.close  
              
              select case chkdate 
                     case xdate

                          if lnflag = "Y" then
                              set ms = conn.execute("select * from loan where memno='"&memno&"' and ldate = '"&lndate&"'  and code='D0' " )
                              if not ms.eof then
                                 pint1  = ms("amount") * 0.01
                                   
                              end if
                              ms.close
                             
                              pint1 = pint1 + round(chequeamt * 0.01 * (mdd - xdd + 1)/mdd,2)
                          else
                              pint1 = round( bal * 0.01 * (mdd - xdd + 1)/mdd,2)                          
                          end if                 
                                        
                     case ydate
                          pass = 0 
                          sqlstr  = "select * from loan where memno='"&memno&"' and ldate>='"&lndate&"'   order by memno,ldate,code  "
                          Ls.open sqlstr, conn, 2, 2 
                          do while  not ls.eof 
                             select case Ls("code")
                                    case "D0"
                                         xint = ls("amount") * 0.01
 
                                    case "F1","MF"
                                      
                                         if ls("amount") = xint then
                                            pass = 1
                                         else
                                         if ls("amount") > xint then
                                            pass = 2
                                         end if
                                         end if
                                 
                              end select 
                          ls.movenext        
                          loop
                          ls.close
                         
                          select case pass  
                                 case 1 
                                      pint1 = round(bal*0.01,2) + round( chequeamt * 0.01 * (mdd - xdd + 1)/mdd,2) 
                                 case 2  
                                      pint1 = round(bal*0.01,2)   
                                 case 0 
                                     if lnflag="Y" then
                                        pint1 = round(bal*0.01,2) + round( chequeamt * 0.01 * (mdd - xdd + 1)/mdd,2) 
                                     else
                                        pint1 = round(bal*0.01,2) + round( bal * 0.01 * (mdd - xdd + 1)/mdd,2) 
                                      end if
                              
                                                   
                          end select  
                      
                     case else              
                         if pint2 > 0 then
                            pint1 = pint2
                         else 
                            pint1 =round( bal * 0.01,2) 
                         end if
              end select

               ttlpamt = pamt1 + pamt2
               ttlpint = pint1 + pint2
               sumttl = ttlpint+ttlpamt
                       samt = 0
                        set ms = server.createobject("ADODB.Recordset")
                        sqlstring= "select sum(amount) from share where memno='"&memno&"' and code='AI' and pflag=1  group by memno  " 
                        ms.open sqlstring,conn,3,3       
                        if not ms.eof then           
                           samt = ms(0)                                       
                        else
                          samt = 0
                        end if  
                        ms.close
               select case mstatus
                        case "A"
                           if sumttl > 0 then
                               if rs("bnklmt") = 0  then
                                  bnklmt = 100000
                               else
                                  bnklmt = rs("bnklmt")
                               end if   
                               if bnklmt > ttlpint  then                                 
                                  conn.execute("insert into autopay (memno,adate,lnnum,code,bankin,curamt,updamt,status,flag,pdate ) values ('"&memno&"','"&mdate&"','"&lnnum&"','F1',"&ttlpint&","&pint1&","&pint2&",'A','N','"&pdate&"' ) ")                                                    
                                   
                                  difamt = bnklmt - ttlpint 
                                  if difamt > ttlpamt then  
                                     conn.execute("insert into autopay (memno,adate,lnnum,code,bankin,curamt,updamt,status,flag,pdate ) values ('"&memno&"','"&mdate&"','"&lnnum&"','E1',"&ttlpamt&","&pamt1&","&pamt2&",'A','N','"&pdate&"' ) ")         
                                     adifamt = difamt - ttlpamt
                                     if adifamt > rs("monthsave") then
                                        if  rs("monthsave")>0 then 
                                           conn.execute("insert into autopay (memno,adate,code,bankin,curamt,updamt,status,flag,pdate) values ('"&memno&"','"&mdate&"','A1',"&rs("monthsave")&","&rs("monthsave")&","&samt&",'A','N','"&pdate&"'  ) ") 
                                        end if
                                     else
                                        conn.execute("insert into autopay (memno,adate,code,bankin,curamt,updamt,status,flag,pdate) values ('"&memno&"','"&mdate&"','A1',"&adifamt&",0,"&samt&",'A','N','"&pdate&"'  ) ") 
                                     end if
                                  else
                                      conn.execute("insert into autopay (memno,adate,lnnum,code,bankin,curamt,updamt,status,flag,pdate ) values ('"&memno&"','"&mdate&"','"&lnnum&"','E1',"&difamt&","&pamt1&","&pamt2&",'A','N','"&pdate&"' ) ")                                                         
                                  end if        

                               else
                                   conn.execute("insert into autopay (memno,adate,lnnum,code,bankin,curamt,updamt,status,flag,pdate ) values ('"&memno&"','"&mdate&"','"&lnnum&"','F1',"&rs("bnklmt")&","&pint1&","&pint2&",'A','N','"&pdate&"' ) ")                                                         
                               end if
                           else
                                  
                                    if samt = 0 then
                                       if rs("monthsave")>0 then                                            
                                          conn.execute("insert into autopay (memno,adate,code,bankin,curamt,updamt,status,flag,pdate) values ('"&memno&"','"&mdate&"','A1',"&rs("monthsave")&","&rs("monthsave")&","&samt&",'A','N','"&pdate&"'  ) ") 
                                       end if
                                     else
                                        ttlsave = samt + rs("monthsave")  
                                        conn.execute("insert into autopay (memno,adate,code,bankin,curamt,updamt,status,flag,pdate) values ('"&memno&"','"&mdate&"','A1',"&ttlsave&","&rs("monthsave")&","&samt&",'A','N','"&pdate&"'  ) ") 
                                     end if              
                             
                           end if    
                                  
                        case "0" 

                             ttlsave = samt + rs("monthsave")
                             if ttlsave > 0 then  
                                conn.execute("insert into autopay (memno,adate,code,bankin,curamt,updamt,status,flag,pdate) values ('"&memno&"','"&mdate&"','A1',"&ttlsave&","&rs("monthsave")&","&samt&",'0','N','"&pdate&"'  ) ") 
                             end if 
                    
                        case "1" 
                             if delyflag = "Y" then
                                conn.execute("insert into autopay (memno,adate,lnnum,code,bankin,curamt,updamt,status,flag,pdate,delyflag ) values ('"&memno&"','"&mdate&"','"&lnnum&"','F1',"&ttlpint&","&pint1&","&pint2&",'1','N','"&pdate&"','"&delyflag&"'  ) ") 
                             else
                                conn.execute("insert into autopay (memno,adate,lnnum,code,bankin,curamt,updamt,status,flag,pdate ) values ('"&memno&"','"&mdate&"','"&lnnum&"','F1',"&ttlpint&","&pint1&","&pint2&",'1','N','"&pdate&"'  ) ") 
                             end if
                             ttlsave = samt + rs("monthsave")
                             if ttlsave > 0 then  
                                if delyflag = "Y" then
                                   conn.execute("insert into autopay (memno,adate,code,bankin,curamt,updamt,status,flag,pdate,delyflag ) values ('"&memno&"','"&mdate&"','A1',"&ttlsave&","&rs("monthsave")&","&samt&",'1','N','"&pdate&"','"&delyflag&"'  ) ") 
                                else 
                                  conn.execute("insert into autopay (memno,adate,code,bankin,curamt,updamt,status,flag,pdate) values ('"&memno&"','"&mdate&"','A1',"&ttlsave&","&rs("monthsave")&","&samt&",'1','N','"&pdate&"'  ) ") 
                                end if  
                             end if 
                       case "2"                

                             conn.execute("insert into autopay (memno,adate,lnnum,code,bankin,curamt,updamt,status,flag,pdate ) values ('"&memno&"','"&mdate&"','"&lnnum&"','E1',"&ttlpamt&","&pamt1&","&pamt2&",'2','N' ,'"&pdate&"' ) ") 
                             ttlsave = samt + rs("monthsave")
                             if ttlsave > 0 then  
                                conn.execute("insert into autopay (memno,adate,code,bankin,curamt,updamt,status,flag,pdate) values ('"&memno&"','"&mdate&"','A1',"&ttlsave&","&rs("monthsave")&","&samt&",'2','N','"&pdate&"'  ) ") 
                             end if 
                       case "3","8"                
                             conn.execute("insert into autopay (memno,adate,lnnum,code,bankin,curamt,updamt,status,flag,pdate ) values ('"&memno&"','"&mdate&"','"&lnnum&"','F1',"&ttlpint&","&pint1&","&pint2&",'3','N','"&pdate&"'  ) ") 
                             conn.execute("insert into autopay (memno,adate,lnnum,code,bankin,curamt,updamt,status,flag,pdate ) values ('"&memno&"','"&mdate&"','"&lnnum&"','E1',"&ttlpamt&","&pamt1&","&pamt2&",'3','N','"&pdate&"'  ) ") 
  
                        case "M"
                              if rs("tpayamt") > ttlpint then
                                 conn.execute("insert into autopay (memno,adate,lnnum,code,bankin,curamt,updamt,status,flag,pdate ) values ('"&memno&"','"&mdate&"','"&lnnum&"','F2',"&ttlpint&","&pint1&","&pint2&",'M','N','"&pdate&"'  ) ")   
                                 difamt = rs("tpayamt") - ttlpint
                                 if difamt < ttlpamt then
                                    conn.execute("insert into autopay (memno,adate,lnnum,code,bankin,curamt,updamt,status,flag,pdate ) values ('"&memno&"','"&mdate&"','"&lnnum&"','E2',"&difamt&","&pamt1&","&pamt2&",'M','N','"&pdate&"'  ) ")       
                                    xdifamt = ttlpamt - difamt 
                                    conn.execute("insert into autopay (memno,adate,lnnum,code,bankin,curamt,updamt,status,flag,pdate ) values ('"&memno&"','"&mdate&"','"&lnnum&"','E1',"&xdifamt&","&pamt1&","&pamt2&",'M','N','"&pdate&"'  ) ")   
                                    conn.execute("insert into autopay (memno,adate,lnnum,code,bankin,curamt,updamt,status,flag,pdate ) values ('"&memno&"','"&mdate&"','"&lnnum&"','A1',"&rs("monthssave")&","&rs("monthssave")&",0,'M','N','"&pdate&"'  ) ")   
                                else
                                    xdifamt = difamt - ttlpamt 
                                    conn.execute("insert into autopay (memno,adate,lnnum,code,bankin,curamt,updamt,status,flag,pdate ) values ('"&memno&"','"&mdate&"','"&lnnum&"','E1',"&ttlpamt&","&pamt1&","&pamt2&",'M','N','"&pdate&"'  ) ")                                       
                                    conn.execute("insert into autopay (memno,adate,code,bankin,curamt,updamt,status,flag,pdate) values ('"&memno&"','"&mdate&"','A1',"&rs("monthssave")&","&rs("monthssave")&",0,'T','N' ,'"&pdate&"' ) ")  
                               
                                end if
                               else
                                   conn.execute("insert into autopay (memno,adate,lnnum,code,bankin,curamt,updamt,status,flag,pdate ) values ('"&memno&"','"&mdate&"','"&lnnum&"','F2',"&rs("tpayamt")&","&pint1&","&pint2&",'M','N','"&pdate&"'  ) ")    
                                   difamt = ttlpint - rs("tpayamt")
                                   conn.execute("insert into autopay (memno,adate,lnnum,code,bankin,curamt,updamt,status,flag,pdate ) values ('"&memno&"','"&mdate&"','"&lnnum&"','F1',"&difamt&","&pint1&","&pint2&",'M','N','"&pdate&"'  ) ")   
                                   conn.execute("insert into autopay (memno,adate,lnnum,code,bankin,curamt,updamt,status,flag,pdate ) values ('"&memno&"','"&mdate&"','"&lnnum&"','E1',"&ttlpamt&","&pamt1&","&pamt2&",'M','N','"&pdate&"'  ) ")    
                                   conn.execute("insert into autopay (memno,adate,lnnum,code,bankin,curamt,updamt,status,flag,pdate ) values ('"&memno&"','"&mdate&"','"&lnnum&"','A1',"&rs("monthssave")&","&rs("monthssave")&",0,'M','N','"&pdate&"'  ) ")   
                               end if





                        case "T" 
                          
                             if rs("tpayamt") > (ttlpint+ttlpamt) then
                                saveamt = rs("tpayamt") - ttlpint - ttlpamt
                                conn.execute("insert into autopay (memno,adate,lnnum,code,bankin,curamt,updamt,status,flag,pdate ) values ('"&memno&"','"&mdate&"','"&lnnum&"','F2',"&ttlpint&","&pint1&","&pint2&",'T','N' ,'"&pdate&"' ) ") 
                                conn.execute("insert into autopay (memno,adate,lnnum,code,bankin,curamt,updamt,status,flag,pdate ) values ('"&memno&"','"&mdate&"','"&lnnum&"','E2',"&ttlpamt&","&pamt1&","&pamt2&",'T','N','"&pdate&"'  ) ")  
                                 
                                if saveamt  > 0 then
                                   
                                   conn.execute("insert into autopay (memno,adate,code,bankin,curamt,updamt,status,flag,pdate) values ('"&memno&"','"&mdate&"','A2',"&saveamt&",0,0,'T','N' ,'"&pdate&"' ) ")  
                                end if
                             else
                                difamt = rs("tpayamt") - ttlpint
				conn.execute("insert into autopay (memno,adate,lnnum,code,bankin,curamt,updamt,status,flag,pdate) values ('"&memno&"','"&mdate&"','"&lnnum&"','F2',"&ttlpint&","&pint1&"  ,"&pint2&",'T','N' ,'"&pdate&"' ) ") 
                                conn.execute("insert into autopay (memno,adate,lnnum,code,bankin,curamt,updamt,status,flag,pdate ) values ('"&memno&"','"&mdate&"','"&lnnum&"','E2',"&difamt&" ,"&pamt1&" , "&pamt2&" ,'T','N' ,'"&pdate&"' ) ")  
                             end if
                   end select     
           else
             select case mstatus 
                    case "A" 
                        if rs("monthsave") > 0 then 
                          conn.execute("insert into autopay (memno,adate,code,bankin,curamt,updamt,status,flag,pdate) values ('"&memno&"','"&mdate&"','A1',"&rs("monthsave")&","&rs("monthsave")&",0,'A','N' ,'"&pdate&"' ) ")  
                        end if
                    case  "M","T" 
                        if rs("monthssave") > 0 then   
                           conn.execute("insert into autopay (memno,adate,code,bankin,curamt,updamt,status,flag,pdate) values ('"&memno&"','"&mdate&"','A2',"&rs("monthssave")&","&rs("monthssave")&",0,'T','N' ,'"&pdate&"' ) ")  
                        end if
              end select
           end if
           
           rs.movenext
        loop   
        rs.close
        conn.committrans           
        msg = "轉賬建立巳完成"
    '    response.redirect "completed.asp" 

end if 

%>
<html>

<head>
<title>轉賬建立</title>

<script language="JavaScript">
<!--
function checkDay(mDay){
  D=mDay.value;
  M=<%=m%>;
  Y=<%=y%>;
  if(isNaN(D) || D=="")
    return false;
  Months= new Array(31,28,31,30,31,30,31,31,30,31,30,31);
  Leap  = false;
  if((Y % 4 == 0) && ((Y % 100 != 0) || (Y %400 == 0)))
    Leap = true;
  if((D < 1) || (D > 31))
    return false;
  if((D > Months[M-1]) && !((M == 2) && (D > 28)))
    return false;
  if(!(Leap) && (M == 2) && (D > 28))
    return false;
  if((Leap)  && (M == 2) && (D > 29))
    return false;
  return true;
}

function validating(){
	formObj=document.form1;
	reqField="";
	placeFocus=false;

	if (formObj.mDay.value==""){
		reqField=reqField+", 日期";
		if (!placeFocus)
			placeFocus=formObj.mDay;
	}

    if (reqField){
        if (reqField.lastIndexOf(",")==0)
	        reqField = "請填入"+reqField.substring(2);
        else
	        reqField = "請填入"+reqField.substring(2,reqField.lastIndexOf(","))+'及'+reqField.substring(reqField.lastIndexOf(",")+2);
        alert(reqField);
        placeFocus.focus();
        return false;
    }else{
        return true;
    }
}
//  -->
</script>
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="form1.mDay.focus()">
<!-- #include file="menu.asp" -->
<br>
<center>
<h3>轉帳建立</h3>
<form name="form1" method="post" action="<% =Request.servervariables("script_name") %>" onsubmit="return validating()">
<%if msg<>"" then %>
<div><center><font color="red"><%=msg%></font></center></div>
<br>
<% end if%>

<table border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td class="b8">日期</td>
		<td width="10"></td>
		<td>
			<input type="text" name="mDay" value="<%=mDay%>" size="2" maxlength="2" onblur="if(!checkDay(this)){this.value=''};">/<%=m%>/<%=y%>
			<input type="submit" value="確定" name="submit" class="sbttn">
		</td>
	</tr>
</table>
</form>
</center>
</body>
</html>
