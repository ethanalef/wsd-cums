<!-- #include file="../conn.asp" -->
<!-- #include file="init.asp" -->
<%


dim sdate(50)
dim scode(50)
dim samt(50)
dim sbal(50)
dim lnnum(50)
dim lndate(50)
dim lcode(50)
dim xcode(50)
dim lnramt(50)
dim lniamt(50)
dim lnbal(50)
dim lncode(50)
dim ldate(50)


server.scripttimeout = 1800
xname = request.form("accode")
pos = instr(xname,"-")

if pos>0 then
memno = mid(xname,1,4)
mname =  mid(xname,pos+1,50)
 
else
  response.redirect "hyPprt.asp"
end if

if memno<>"9999"  then
memno = left(xname,pos-1)
mname =  mid(xname,pos+1,50)
styfield = "accode='"&memno&"' and "
else
if instr(mname,"(1-1500)") > 0 then
             styfield = "memno<='1500' and "
else
if instr(mname,"(1501-3000)") > 0 then
      styfield = " (memno>='1501' and  memno<='3000') and "
else 
if instr(mname,"(3001-5000)") > 0 then
      styfield = " memno >'3000' and  "  
else
   styfield = ""
end if
end if
end if
end if

yy =request.form("Nyear")
Nyear = (yy-1)&"/"&yy

xxdate =dateserial(yy-1,7,1)
yydate =dateserial(yy,7,1)
chkdate ="01/07/"&(yy-1)
ndate=dateserial(2008,4,30)
if xxdate < ndate then
   xxdate = ndate
   chkdate ="30/04/"&(yy-1)
end if

if memno ="9999" then
   xmemname = ""
   xmemcname = "工作人員"
   xmemContactTel = "27879222"

else
set rs=conn.execute("select memno,memname,memcname,memContactTel from memmaster where memno='"&memno&"' ")
if not rs.eof then
   xmemname = rs("memname")
   xmemcname = rs("memcname")
   xmemContactTel = rs("memContactTel")
end if
rs.close

end if
dim divd(12)
dim dvdate(12)
xyear = yy - 1
for i = 1 to 12
    dvdate(i)=dateserial(xyear,6+i,1)

next  
SQl = "select memno,memname,memcname  from memmaster where  accode='"&memno&"'   and mstatus not in ('8','9','C','P','V','L','B' ) order by memno   "
Set rs = Server.CreateObject("ADODB.Recordset")
Set ms = Server.CreateObject("ADODB.Recordset")
Set ns = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn,1,1

if request.form("output")="Word" then
	Response.ContentType = "application/msword"
elseif request.form("output")="excel" then
	Response.ContentType = "application/vnd.ms-excel"

end if


%>
<html>
<head>
<title>全年結</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">



<%
   do while not rs.eof
      xmemno = rs("memno")
      memcname = rs("memcname")
      memname = rs("memname")
      line = 8

     for i = 1 to 50
         sdate(i) = ""
         scode(i) = "" 
         samt(i)  = ""
         sbal(i)  = ""
         lnnum(i) = ""
         lndate(i) = ""
         lncode(i)  = ""
         xcode(i)  = ""
         lnramt(i) = ""
         lniamt(i) = ""
         lnbal(i)  = ""
        
         ldate(i)  = ""
         lcode(i) = "" 
      next  
       xlnnum=""
       ylnnum = ""

 set ms=conn.execute("select lnnum from loan where memno='"&xmemno&"' and ldate>='"&xxdate&"' and ldate<='"&yydate&"' order by memno,ldate desc ,code ")
if not ms.eof then
   xlnnum = ms("lnnum")
 

end if
ms.close

SQl = "select lnnum,ldate  from loan  where memno='"&xmemno&"' and ldate>='"&xxdate&"' and ldate<='"&yydate&"' order by lnnum,ldate,right(code,2),left(code,1)  " 
set rs1 = Server.CreateObject("ADODB.Recordset")
rs1.open sql, conn,1,1
xx = 0
if not rs1.eof then
   ylnnum = rs1(0)
   xldate = rs1(1)
end if
rs1.close

       
       xx = 1
       yy = 0
       cut = 1
       if ylnnum<>"" then
          set ms = conn.execute("select *  from loan where  memno='"&xmemno&"'  and ldate<='"&yydate&"'  order by memno,ldate,right(code,1),left(code,1) ") 
             cc = 0
             xx = 1
             qx = 0 
             MX=0
             zero = 0
             bb = 0
             yy = 0
             lnbal(xx) = 0
          do while not ms.eof
             ssdate = right("0"&day(ms("ldate")),2)&"/"&right("0"&month(ms("ldate")),2)&"/"&year(ms("ldate"))
             if ms("ldate")>xxdate then
        	  if lnbal(xx) <> "" and  yy = 0 then
                     if lnbal(xx) > 0 then
                        bal = lnbal(xx)
                          xx = xx + 1	
                     else
                        lnbal(xx) = ""
                        bal = 0 
                     end if
            	  
                    yy = 1
           
        	 end if      	 
             xdate = right("0"&day(ms("ldate")),2)&"/"&right("0"&month(ms("ldate")),2)&"/"&year(ms("ldate"))
             select case ms("code")
           
          case "0D" 
   		lnbal(xx)=ms("amount")
        	bal = ms("amount")
                lnbal(xx)=bal
         case  "D8"
                lnramt(xx) =ms("amount")
                bal = bal - ms("amount")
                lnbal(xx) = bal
          case "D9"
                set ms1 =  conn.execute("select appamt from loanrec where memno = '"&ms("memno")&"' and lndate= '"&ms("ldate")&"'   ")
                bal= lnamt + ms1(0)
                lnbal(xx)=bal
                ms1.close 
          case  "E0","E1" , "E4" , "E2" , "E3" , "E6" ,"E7","EC"
                bal = bal - ms("amount")
                lnbal(xx) = bal
          case  "ER"
               bal = bal + ms("amount")
                lnbal(xx) = bal                
          case  "DE" 
                lnbal(xx) = bal
       
   end select 

   if ms("code")="F3" and  xdate <> lndate(xx-1) then

      mx = 0
   end if

   if left(ms("code"),1) ="E" or ms("code")="0D" or ms("code")="D8" or ms("code")="D9" OR ms("code")="DE" or ms("code")="NE" or  ((ms("code")="F3"  or ms("code")="F1" or ms("code")="F2" or ms("code")="F0" ) and MX=0)  then
      lnnum(xx) = ms("lnnum")
      newln = 0
      xyear = year(ms("ldate"))
      xmon  = month(ms("ldate"))
      xday  = day(ms("ldate"))
      oldamt = 0
      xcode(xx) = ms("code")
      ldate(xx) = ms("ldate")
      lndate(xx) =right("0"&xday,2)&"/"&right("0"&xmon,2)&"/"&xyear
      xdate      =right("0"&xday,2)&"/"&right("0"&xmon,2)&"/"&xyear     
      lcode(xx) = LoanCode(ms("code")) 
   
     if (ms("CODE")="F3" or ms("code")="F1" or ms("code")="F2"  or ms("code")="F0") AND mx = 0 THEN
         LNIAMT(XX)=ms("AMOUNT")
         MX = 0 
     ELSE  
         if left(ms("code"),1)="E" OR ms("code")="D8" or  ms("code")="0D" or  ms("code")="DE" or ms("code")="NE" then
       
            if ms("amount") <> 0 then 
               if ms("code")="0D" then
                   lnbal(xx) = ms("amount")
               else
                  lnramt(xx) =ms("amount")
               end if
            end if         
         else
            if ms("code")="D9"  then            
               set ms1=conn.execute("select chequeamt,lnflag,appamt,loantype from loanrec where lnnum='"&ms("lnnum")&"'  ")
               if  not ms1.eof then
                   loantype = ms1("loantype")   
                             
                   lnflag = ms1(1)
                   if lnflag = "Y" then                        
                      
                      if ms1(0) <> 0 then  
                         lniamt(xx) =  ms1(0)
                      else
                         lniamt(xx) =""
                      end if
                      if loantype ="N" then
                        
                         lcode(xx) = "+ 新貸  ="
                      else
                     
                      lcode(xx)=" 更改期數 "
                   end if               
                      
                         lnbal(xx) = ms1(2)  
                    
                      bal = lnbal(xx)          
                    ELSE
                      lnbal(xx) = ms1("appamt")
                      lniamt(xx) =""
                      bal = lnbal(xx)
                      lcode(xx) ="新貸"  
                    END IF                     
                 end if
                ms1.close 
                MX = 0     
             end if 
           END IF
        end if
    
    xx = xx + 1
    bb = xx
    if ms("code")="D8" and ms("amount")= 0 then
       xx = xx -  1
       zero= 1
    end if
   end if 
   if  ms("code")="DF" or ms("code")="NF"   then
       mx = 0
       lniamt(xx-1) = ms("amount")    
   end if  

  if left(ms("code"),1) ="E"   then
     mx = mx + 1
     tmpdate =  ms("ldate") 
    
        ms.movenext
    
   
     if not ms.eof then    
        select case   left(ms("code"),1)
            case "E"
                 if ms("ldate") <> tmpdate then
                    mx = 0   
 
                 end if
                           
             case "F"
                              
                     if mx = 1 then
                         lniamt(xx-1)= ms("amount") 
                         mx = 0
                     else
                       select case ms("code")
                              case "F0","F3"
                                    lniamt(xx-2)= ms("amount")
                                    if  not ms.eof then
                                    ms.movenext
                                    if left(ms("code"),1)="F" then
                                       lniamt(xx-1)= ms("amount")
                                    else
                                       ms.moveprevious  
                                    end if 
                                    end if
                              case "F1"
 
                                              if right(lcode(xx-2),1) = left(ms("code"),1) then
                                                 lniamt(xx-2)= ms("amount")
                                              else
                                                 lniamt(xx-1)= ms("amount")
                                             end if 
 
                                    if  not ms.eof then
                                    ms.movenext
                                    if not ms.eof then                                   
                                     if left(ms("code"),1)="F" then
                                        lniamt(xx-1)= ms("amount")
                                     else
                                        ms.moveprevious  
                                    end if  
                                    end if
                                    end if                 
                              case "F2"
                                   lniamt(xx-1)= ms("amount") 
                                                  
                        end select             
                        mx = 0

                    end if
                         
         case else
                mx = 0
                ms.moveprevious  

         end select 
         
    end if
  end if
  else

        if  left(ms("code"),1)="E" or ms("code")="D8"  then        
               lnbal(xx)=lnbal(xx)-ms("amount")
        else 
        if ms("code")="0D" or ms("code")="D9"   then 
         
            lnbal(xx) = lnbal(xx) +ms("amount")
           end if
        end if  
           
  
            lndate(xx) = ssdate
            lcode(xx) = "貸款結餘" 
            lniamt(xx)  =""     
            aa = 1
end if
  

          

if  not ms.eof then
ms.movenext 
end if
loop
ms.close

end if
      bb = xx
       
      yy = 0
      xx = 1
      aa = 0
    set ms = conn.execute("select * from share where memno='"&xmemno&"' and ldate<='"&yydate&"'  order by memno,ldate,code ")
     if not ms.eof then
       for i = 1 to 50
          sbal(i) = 0
      next 
      for i = 1 to 12
          divd(i)=0
      next 
      do while not ms.eof 
         xyear = year(ms("ldate"))
         xmon  = month(ms("ldate"))
         xday  = day(ms("ldate"))
         xdate = xyear&xmon&xday

        ssdate = right("0"&xday,2)&"/"&right("0"&xmon,2)&"/"&xyear 
         if ms("ldate") > xxdate then
 
 	  if sbal(xx) > 0 and  yy = 0 then
            divd(1) = sbal(xx)
  	    xx = xx + 1	
            yy = 1
	 end if
    select case ms("code")
        case "0A"
                sbal(xx)=sbal(xx-1)+ms("amount")
                   mkdate1 = dateserial(year(ms("ldate")),month(ms("ldate"))+1,1)
                   for i = 1 to 12
                      if dvdate(i)= mkdate1 then
                       
                          divd(i) = sbal(xx)
                       end if 
                      
                 
                next 
        case "G0" ,"H0","B0","B1","B3","BE","BF","G3","H3","MF" 
               sbal(xx)=sbal(xx-1)-ms("amount")
               if left(ms("code"),1)="B"  then
                  mkdate2 = dateserial(year(ms("ldate")),month(ms("ldate")),1)
                  if mkdate0 <= ms("ldate") then
                     for i = 1 to 12
                        if dvdate(i)= mkdate1 then
                           divd(i) = sbal(xx)
                        end if
                     next                      
                  end if
                  for i = 1 to 12
                       if dvdate(i)<= mkdate2 then
                          divd(i) = sbal(xx)
                       end if
                  next      
               end if   
         case "C5"
              sbal(xx)=sbal(xx-1)+ms("amount")
              response.write(sbal(xx))
                response.end
                  mkdate2 = dateserial(year(ms("ldate")),month(ms("ldate")),1)
                  if mkdate0 <= ms("ldate") then
                     for i = 1 to 12
                        if dvdate(i)= mkdate1 then
                           divd(i) = sbal(xx)
                        end if
                     next                      
                  end if
                  for i = 1 to 12
                       if dvdate(i)<= mkdate2 then
                          divd(i) = sbal(xx)
                       end if
                  next      
                
        case   "AI","CH"

                  sbal(xx) = sbal(xx-1)  
                   mkdate1 = dateserial(year(ms("ldate")),month(ms("ldate"))+1,1)
                   for i = 1 to 12
                      if dvdate(i)= mkdate1 then
                       
                          divd(i) = sbal(xx)
                       end if 
                      
                 
                next 
   
        case  "A1","A2","A3","C0","C1","C3" ,"A0","A7" ,"A4" 
                   if ms("code") = "C5" then
                         response.write("C5")
                           response.end
                     end if
                sbal(xx) = sbal(xx-1) + ms("amount")

                  
                mkdate1 = dateserial(year(ms("ldate")),month(ms("ldate"))+1,1)
                mkdate0  = ms("ldate")
                for i = 1 to 12
                    if dvdate(i)= mkdate1 then
                     
                          divd(i) = sbal(xx)
                       end if 
                      
                 
                next 

        end select
        scode(xx) = shareCode(ms("code"))

        sdate(xx) = ssdate
    if ms("amount") <> "" then  
        if ms("code")<>"0A" then     
           samt(xx) =ms("amount")
        end if
    end if
        xx = xx + 1
        aa = xx
         else
           
                if  left(ms("code"),1)="G" or left(ms("code"),1)="H" or left(ms("code"),1)="B" then        
                    sbal(xx)=sbal(xx)-ms("amount")
                else 
                if ms("code")="0A" or left(ms("code"),1)="C" or left(ms("code"),1)="A" then 
                   sbal(xx) = sbal(xx) +ms("amount")
                end if  
	       
                sdate(xx) = chkdate
                scode(xx) = "股金結餘" 
                samt(xx)  =""     
                aa = 1
               end if
           end if
        ms.movenext
        loop
        end if
        ms.close 
      
xx = 1


if aa > bb then
   smax = aa
else
   smax = bb
end if
if sbal(1) = ""  then
   sbal(1) =""
   scode(1)=""
   
end if

%>

<center>
<BR>
<font face="arial, helvetica, sans-serif" size="3" color="#336699"><b>水務署員工儲蓄互助社系統</b></font><br>
<font face="arial, helvetica, sans-serif" size="1" color="#336699"><b>Water Supplies Department Staff Credit Union<br>Membership, Accounting, Savings and Loans Software</b></font>

<BR>
<font face="arial, helvetica, sans-serif" size="2" color="#336699"><b><%=Nyear%>　全年結</b></font>
</center>
<table width="1060" border="0" cellspacing="0" cellpadding="0">
     <tr>
          <td width="400" align="left"><font size="3"  face="標楷體" >　　<%=xmemno%>　<%=memcname%></font></td>
          <td width="660" align="left">　　</td>
     </tr>
     <tr>
          <td width="400" align="left">　　 <%=memname%></font></td> 
          <td width="660" align="left"><font size="3"  face="標楷體" >　　　　（聯絡員）<%=xmemcname%>（電話）<%=xmemContactTel%></font></td> 
     </tr>
     
</table>


<br>
<table border="0" cellspacing="1" cellpadding="1" align="center" >
	<tr >
		
		<td width=100 align="center"><font size="2"  face="標楷體" >日期</font></td> 
               
		<td width=80 align="center"><font size="2"  face="標楷體" >股金</font></td> 
               
		<td width=100 align="center"><font size="2"  face="標楷體" >結餘</font></td> 
                            
		<td width=100 align="center"><font size="2"  face="標楷體" >類別</font></td> 
               
                <td  width=1 align="center"> </td>	
		<td width=100 align="center"><font size="2"  face="標楷體" >日期</font></td> 
               

                
		<td width=100 align="center"><font size="2"  face="標楷體" >利息</font></td> 
               
		<td width=100 align="center"><font size="2"  face="標楷體" >每月還款</font></td> 
                
		<td width=100 align="center"><font size="2"  face="標楷體" >新貸總額/<br>結餘</font></td> 
              
		<td width=100 align="center"><font size="2"  face="標楷體" >類別</font></td>            
	</tr>
	
	<tr><td colspan=11><hr></td></tr>
<%
       xx = 1
       do while xx < smax

%>  
        <tr>
		<td width=100 align="center"><%=sdate(xx)%></td>
<%if samt(xx)<>"" then%>               
		<td width=80 align="right"><%=formatnumber(samt(xx),2)%></td>
<%else%>
                <td width=80 align="right"></td>
<%end if%>
<%if sbal(xx)<>"" then
  if sbal(xx)>0 then %> 
                

		<td width=100 align="right"><%=formatnumber(sbal(xx),2)%></td>
<%else%>
                <td width=100 align="right">&nbsp;</td>
<%end if%>
<%end if%>                            
		<td width=100 align="center"><font size="2"  face="標楷體" ><%=scode(xx)%></font></td>    
               <td  width=1 align="center"> </td>	

               
		<td width=100 align="center"><%=lndate(xx)%></td>
<%if lniamt(xx)<>"" then%>                
		<td width=100 align="right"><%=formatnumber(lniamt(xx),2)%></td>
<%else%>
               <td></td>
<%end if%>
<%if lnramt(xx)<>"" then%>                
		<td width=100 align="right"><%=formatnumber(lnramt(xx),2)%></td>
<%else%>
                <td></td>
<%end if%>
<%if lnbal(xx)<>""   then%>                 

		<td width=100 align="right"><%=formatnumber(lnbal(xx),2)%></td>

<%else%>
                <td>&nbsp;</td>
<%end if%>
              
		<td width=100 align="center"><font size="2"  face="標楷體" ><%=lcode(xx)%></font></td>         
        </tr>   
  
<%
        xx = xx + 1
        line = line + 1
        loop
      if smax< 18 then
         uu = 18 - smax 
         for i = 1 to uu
            
%>
          <tr><td colspan=11>&nbsp;</td></tr>
<%
      line = line + 1
      next
      end if   

%>
  </table>

<%      


if xlnnum<>"" then
%>        

<table border="0" cellspacing="1" cellpadding="1" align="center" >

   <tr>
         <td width="100" >借據編號</td>
         <%if xlnnum<>"" then %>
         <td width="100" ><%=xlnnum%></td> 
         <%else%>
         <td width="100" >&nbsp;</td>
         <%end if%>
         <td width="100" >&nbsp;</td>
    </tr> 
    </table>
<%
line = line   + 1

sql1 = "select * from guarantor  where  guarantorID = "&xlnnum
ms.open sql1, conn, 1, 1
do while  not ms.eof 
     select case xx
             case 1 
                  guoid1 = rs("memno")                              
                  if guoid1 <> "" then
                  guln1=""
                  Set rs1 = server.createobject("ADODB.Recordset")
		  sql1 = " select * from loanrec where repaystat='N' and memno = "& guoid1                  
                  rs1.open sql1, conn, 2, 2
                  if  not rs1.eof then
                      guln1 = rs1("lnnum")
                  end if 
                             
                  
                  rs1.close
                  Set rs1 = server.createobject("ADODB.Recordset")
		  sql1 = " select memname,memcname from memmaster where memno = "& guoid1                  
                  rs1.open sql1, conn, 2, 2
                  if  not rs1.eof then
                      guoname1 = rs1("memname")&" "&rs1("memcname")
                  end if 
                             
                  
                  rs1.close                
                 end  if
             case 2
		 guid2 = rs("guarantorID")               
                 guname2 = rs("guarantorName")
                 if guoid2 <> "" then
                  guln2=""
                  Set rs1 = server.createobject("ADODB.Recordset")
		  sql1 = " select * from loanrec where repaystat='N' and memno = "& guoid2                  
                  rs1.open sql1, conn, 2, 2
                  if  not rs1.eof then
                      guln2 = rs1("lnnum")&" "&rs1("memcname")
                  end if 
                  
                  rs1.close
                
                  Set rs1 = server.createobject("ADODB.Recordset")
		  sql1 = " select memname,memcname from memmaster where memno = "& guoid2                  
                  rs1.open sql1, conn, 2, 2
                  if  not rs1.eof then
                      guoname2 = rs1("memname")&" "&rs1("memcname")
                  end if 
                             
                  
                  rs1.close  
                end if
             case 3
		 guid3 = rs("guarantorID")            
                 guname3 = rs("guarantorName")
                 if guoid3 <> "" then
                  guln3=""
                  Set rs1 = server.createobject("ADODB.Recordset")
		  sql1 = " select * from loanrec where repaystat='N' and memno = "& guoid3                  
                  rs1.open sql1, conn, 2, 2
                  if  not rs1.eof then
                    
                          guln3 = rs1("lnnum")
                  end if 
                             
                  
                  rs1.close
                  Set rs1 = server.createobject("ADODB.Recordset")
		  sql1 = " select memname,memcname from memmaster where memno = "& guoid13                 
                  rs1.open sql1, conn, 2, 2
                  if  not rs1.eof then
                      guoname3 = rs1("memname")+" "+rs1("memcname")
                  end if 
                             
                  
                  rs1.close                   
                  end if
         end select 

ms.movenext
loop
ms.close
set ms=nothing
set rs1=nothing

%>
<table border="0" cellspacing="1" cellpadding="1" align="center"  >
<%if guid1 <>"" then 
   
%>
        <tr>
             <td width="50">&nbsp;</td>
             <td width="100">1.擔保人<td>
             <td width="50"><%=guid1%></td>
             <td width="200"><%=guname1%></td>
             
      
        </tr>
<%else%>
        <tr>
              <td width="50">&nbsp;</td>
             <td width="100">1.擔保人<td>
             <td width="50">*****</td>
             <td width="200">*****</td>
 
      
        </tr>
<%end if%>
<%if guid2 <>"" then 
    
%>
        <tr>
             <td width="50">&nbsp;</td>
             <td width="100">2.擔保人 <td>
             <td width="50"><%=guid2%></td>
             <td width="200"><%=guname2%></td>
 
      
        </tr>
<%else%>
        <tr>
             <td width="50">&nbsp;</td>
             <td width="100">2.擔保人<td>
             <td width="50">*****</td>
             <td width="200">*****</td>
 
      
        </tr>
<%end if%>
<%if guid3 <>"" then
    
%>
        <tr>
          
              <td width="50">&nbsp;</td>
             <td width="100">3.擔保人 <td>
             <td width="50"><%=guid3%></td>
             <td width="200"><%=guname3%></td>
 
      
        </tr>
<%else%>
        <tr>
             <td width="50">&nbsp;</td>
             <td width="100">3.擔保人<td>
             <td width="50">*****</td>
             <td width="200">*****</td>
 
      
        </tr>
<%end if%>
<%if guoid1 <>"" then
     
%>
        <tr>
              <td width="50">&nbsp;</td>
             <td width="100">1.擔保其他人 <td>
             <td width="50"><%=guoid1%></td>
             <td width="300"><%=guoname1%></td>
  
      
        </tr>
<%else%>
        <tr>
              <td width="50">&nbsp;</td>
             <td width="100">1.擔保其他人<td>
             <td width="50">*****</td>
             <td width="200">*****</td>
 
      
        </tr>
<%end if%>
<%if guoid2 <>"" then
     
%>
        <tr>
              <td width="50">&nbsp;</td>
             <td width="100">2.擔保其他人 <td>
             <td width="50"><%=guoid2%></td>
             <td width="200"><%=guoname2%></td>

      
        </tr>
<%else%>
        <tr>
              <td width="50">&nbsp;</td>
             <td width="100">2.擔保其他人<td>
             <td width="50">*****</td>
             <td width="200">*****</td>
 
      
        </tr>
<%end if%>
<%if guoid3 <>"" then 
    
%>
        <tr>
             <td width="50">&nbsp;</td>
             <td width="100">3.擔保其他人 <td>
             <td width="50"><%=guoid3%></td>
             <td width="200"><%=guoname3%></td>
 
        </tr>
<%else%>
        <tr>
              <td width="50">&nbsp;</td>
             <td width="100">3.擔保其他人<td>
             <td width="50">*****</td>
             <td width="200">*****</td>
 
      
        </tr>
<%end if%>
</table>
     
<% 
line = line + 6 
else
%>
<table border="0" cellspacing="1" cellpadding="1"  align="center" >


   <tr>
         <td width="100" >借據編號</td>
 
         <td width="100" >&nbsp;</td>
 
         <td width="100" >&nbsp;</td>
    </tr> 
    </table>
    <table border="0" cellpadding="0" cellspacing="0"  align="center" >
        <tr>
              <td width="50">&nbsp;</td>
             <td width="100"><b>1.擔保人</b><td>
             <td width="50">*****</td>
             <td width="200">*****</td>

      
        </tr>
        <tr>
              <td width="50">&nbsp;</td> 
             <td width="100"><b>2.擔保人</b><td>
             <td width="50">*****</td>
             <td width="200">*****</td>
 
        </tr>
        <tr>
             <td width="50">&nbsp;</td>
             <td width="100"><b>3.擔保人</b><td>
             <td width="50">*****</td>
             <td width="200">*****</td>
 
      
        </tr>
        <tr>
              <td width="50">&nbsp;</td>  
             <td width="100"><b>1.擔保其他人</b><td>
             <td width="50">*****</td>
             <td width="200">*****</td>

      
        </tr>
        <tr>
              <td width="50">&nbsp;</td>
             <td width="100"><b>2.擔保其他人</b><td>
             <td width="50">*****</td>
             <td width="200">*****</td>
 
      
        </tr>
        <tr>
             <td width="50">&nbsp;</td>
             <td width="100"><b>3.擔保其他人</b><td>
             <td width="50">*****</td>
             <td width="200">*****</td>
 
        </tr>

  </table>
<%
line = line + 7
end if
      
%>
<center>
<font face="arial, helvetica, sans-serif" size="2" ><b>　　　　股息計算表</b></font><br>
</center>
        <table border="1" cellpadding="0" cellspacing="0" align="center" >

        <tr>
           
            <td width="100" align="center">　　日期　　</td>
            <td width="200" align="center"  >　　內容　　</td>
            <td width="100" align="center" >　　金額　　</td>
        </tr>
<%
        rate = 5 
        ttlamt = 0
        for i = 1 to 12
           sldate = "01/"&right("0"&month(dvdate(i)),2)&"/"&year(dvdate(i))
           if divd(i)= 0 then
          
              divd(i)= divd(i-1)
           end if
           desc   = divd(i)&"X"&rate&"/"&"100"&"/12"
           xamt   =round( divd(i)*rate/100/12,2)
           ttlamt = ttlamt + xamt
%>
        <tr>
            
            <td width="100" align="center"><%=sldate%>前</td>
            <td width="200" align="left"  ><%=desc%></td>
            <td width="100" align="right" ><%=formatnumber(xamt,2)%></td>
        </tr>


<%
        next
%>
        <tr>
            
            <td width="100" align="center">&nbsp;</td>
            <td width="200" align="right"  >合共金額：</td>
            <td width="100" align="right" ><%=formatnumber(ttlamt,2)%></td>
        </tr>

        </table>
<br>


<font face="arial, helvetica, sans-serif" size="2" ><b>　　　　監察委員會</b></font><br>

<font face="arial, helvetica, sans-serif" size="2" ><b>　　　　註 - 協會費由本會支付</b></font><br>
<%if ttlamt > 100 then %>
<font face="arial, helvetica, sans-serif" size="2" ><b>　　　　註 - 本年度股息為<%=formatnumber(ttlamt,2)%>元，已自動轉帳入貴銀行戶口</b></font><br>
<%else%>
<font face="arial, helvetica, sans-serif" size="2" ><b>　　　　註 - 本年度股息為<%=formatnumber(ttlamt,2)%>元，已存入股金</b></font><br>
<%end if%>
<br>
<font face="arial, helvetica, sans-serif" size="2" ><b>　　　　此結餘表字體會色，若需要保存，</b></font><br>
<font face="arial, helvetica, sans-serif" size="2" ><b>　　　　請自行影印</b></font><br>
<br>
<br>
<br>



<table border="0" cellspacing="1" cellpadding="1" align="center" >
<%
     line = line + 12
     uu = 56 - line

        for i = 1 to uu
%>
       <tr>
             <td width="50">&nbsp;</td>
             <td width="100">&nbsp;<td>
             <td width="50">&nbsp;</td>
             <td width="200">&nbsp;</td>
             <td width="100">&nbsp;</td>
             <td width="150"  align="right" >&nbsp;</td>
      
        </tr>
<%
       next
%>
</table>
<%
       for i = 1 to 12
           divd(i) = 0
       next 
       rs.movenext
       loop
%>

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
