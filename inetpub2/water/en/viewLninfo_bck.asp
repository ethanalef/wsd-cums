<!-- #include file="../conn.asp" -->

<%
memno =request("key")

pos = instr(memno,"*")
chkdate=right(memno,10)
id = left(memno,pos-1)

if id="" then
%>
<script language="JavaScript">
<!--

  
	window.opener.document.form1.memNo.focus()        
	
	window.close();
//-->
</script>

<%
    response.end
end if

mndate=right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
xmon = mid(chkdate,4,2)
xyr1  = cint(right(chkdate,4) )
dd    = cint(left(chkdate,2))
xdd    = dd - 1
xlnum = ""
mdate =xyr1&"."&right("0"&xmon,2)&".01"

if xdd = 0 then
    
   xmn = xmon - 1
   if xmn = 0 then
      xmn = 12
      xyr1 = xyr1 - 1
   end if
   if  int(xyr1/4) = xyr1/4 and int(xyr/100)=xyr1/100 then
       xdd = cint(mid("312931303130313130313031",(xmn-1)*2+1,2))
   else
       xdd = cint(mid("312831303130313130313031",(xmn-1)*2+1,2))
   end if

else
   xmn = xmon
end if   

if xmn = 0 then
   xmn = 12
   xyr1 =xyr1-1
end if


xmdate =xyr1&"/"&right("0"&xmon,2)
ydate =xdd&"/"&right("0"&xmn,2)&"/"&xyr1
ndate =xyr1&right("0"&xmn,2)&right("0"&xdd,2)

Set rs = server.createobject("ADODB.Recordset")
sql = "select memno,memname,memcname from memmaster where memno='"&id&"' "
rs.open sql,conn,1,1
if not rs.eof then
   memname = rs("memname")
   memcname= rs("memcname")
end if
rs.close

xlnnum = ""
sql = "select lnnum from loan where ldate>='"&mdate&"'  and  memno = '"& id &"' order by  ldate " 
rs.open sql, conn, 1, 1
if  not rs.eof then
   
      xlnnum=rs("lnnum")
else
set rs1 = conn.execute("select  lnnum,lndate  from loanrec where memno='"&id&"' order by lnnum  desc ")
if not rs1.eof then
   xldate = rs1(1)
   xlnnum = rs1(0)
end if
rs1.close 
   
end if

rs.close

if xlnnum <> "" then
   sql = " select * from guarantor where lnnum = "& xlnnum
   rs.open sql, conn, 1, 1
   xx = 1
   do while  not rs.eof 
      select case xx
             case 1 
                  guid1 = rs("guarantorID")                
                  guname1 = rs("guarantorName")
                  if guid1 <> "" then
                  bal = 0
                  Set rs1 = server.createobject("ADODB.Recordset")
		  sql1 = " select * from share where memno = "& guid1                  
                  rs1.open sql1, conn, 2, 2
                  do while not rs1.eof
                     select case left(rs1("code"),1)
                            case "G","H","B"
                                 bal = bal - rs1("amount")
                            case "A","T","C","0"
                                 bal = bal+ rs1("amount")
                     end select
                             
                  rs1.movenext
                  loop
                  rs1.close
                  gusave1 = bal
                  end if
             case 2
		 guid2 = rs("guarantorID")               
                 guname2 = rs("guarantorName")
                  if guid2 <> "" then
                  bal = 0
                  
		  sql1 = " select * from share where memno = "& guid2                  
                  rs1.open sql1, conn, 2, 2
                  do while not rs1.eof
                     select case left(rs1("code"),1)
                            case "G","H","B"
                                 bal = bal - rs1("amount")
                            case "A","T","C","0"
                                 bal = bal+ rs1("amount")
                     end select
                             
                  rs1.movenext
                  loop
                  rs1.close
                  gusave2 = bal
                  end if
             case 3
		 guid3 = rs("guarantorID")            
                 guname3 = rs("guarantorName")
                  if guid3 <> "" then
                  bal = 0
                 
		  sql1 = " select * from share where memno = "& guid3                  
                  rs1.open sql1, conn, 2, 2
                  do while not rs1.eof
                     select case left(rs1("code"),1)
                            case "G","H","B"
                                 bal = bal - rs1("amount")
                            case "A","T","C","0"
                                 bal = bal+ rs1("amount")
                     end select
                             
                  rs1.movenext
                  loop
                  rs1.close
                  gusave3 = bal
                  end if
         end select 
        
      xx = xx + 1 
   rs.movenext   
   loop
   rs.close
end if
xx = 1

sql = "select * from guarantor  where  guarantorID = "&id
rs.open sql, conn, 1, 1
do while  not rs.eof 
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
        
      xx = xx + 1 
             
rs.movenext
loop
rs.close
set rs=nothing
set rs1=nothing




SQl = "select lnnum,ldate  from loan  where   lnnum='"&xlnnum&"' and code='D1'  order by lnnum,ldate,right(code,2),left(code,1)  " 
set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn,1,1
xx = 0
if not rs.eof then
   xldate = rs(1)
end if
rs.close

dim sdate(500)
dim scode(500)
dim samt(500)
dim sbal(500)
dim lnnum(500)
dim lndate(500)
dim lcode(500)
dim lnramt(500)
dim lniamt(500)
dim lnbal(500)
dim lncode(500)


scode(1) = "股金結餘"


if xlnnum <> "" then
SQl = "select  *,convert(char(10),ldate,102) as Expr1  from loan  where memno='"&id&"' and ldate  >='"&xldate&"' order by memno,ldate,right(code,1),left(code,1) "
Set rs = Server.CreateObject("ADODB.Recordset")	
rs.open sql, conn,2  

cc = 0
xx = 1
qx = 0 = 0
MX=0
zero = 0
    
 
  
do while not rs.eof
  lncode(xx) = rs("code")
  xdate = right("0"&day(rs("ldate")),2)&"/"&right("0"&month(rs("ldate")),2)&"/"&year(rs("ldate"))
   select case rs("code")
           
          case "0D","D1" 
   		lnbal(xx)=0
        	bal = rs("amount")
                lnbal(xx)=""
          case  "E0","E1" , "E4" , "E2" , "E3" , "E6" 
                bal = bal - rs("amount")
                lnbal(xx) = bal
          case  "ER"
               bal = bal + rs("amount")
                lnbal(xx) = bal                
          case  "ME" 
                lnbal(xx) = bal
       
   end select 

   if rs("code")="F3" and  xdate <> lndate(xx-1) then

      mx = 0
   end if

   if left(rs("code"),1) ="E" or rs("code")="0D" or rs("code")="D0" or rs("code")="D1" OR rs("code")="ME" or rs("code")="NE" or  ((rs("code")="F3"  or rs("code")="F1" or rs("code")="F2" ) and MX=0)  then
      lnnum(xx) = rs("lnnum")
      newln = 0
      xyear = year(rs("ldate"))
      xmon  = month(rs("ldate"))
      xday  = day(rs("ldate"))
      oldamt = 0
      lndate(xx) =right("0"&xday,2)&"/"&right("0"&xmon,2)&"/"&xyear
      xdate      =right("0"&xday,2)&"/"&right("0"&xmon,2)&"/"&xyear     
      select case rs("code")
          case "0D"
               lcode(xx) ="貸款結餘"
          case "E1"
               lcode(xx) = "銀行轉帳"
          case "E2"
		 lcode(xx) ="庫房轉帳"
          case "E3"
		 lcode(xx) ="現金還款"
          case "E0"
               if rs("amount") > 0 then
  		   lcode(xx) ="股金還款"
               else
                  lcode(xx) ="退款"
               end if 
                  
          case "E6"
                    lcode(xx) ="退款"

          case "F1"
                 lcode(xx) ="銀行還息"
          case "F2"
                 lcode(xx) ="庫房還息"
          case "F3"
                 lcode(xx) ="現金還息"
          case "ER"
		 lcode(xx) ="退還本金"
          case "F3"
                 lcode(xx) ="現金還息"  
          case "FR"
		 lcode(xx) ="退還利息"
          CASE "ME"
               lcode(xx) ="銀行脫期" 
          CASE "MF"
               lcode(xx) ="利息脫期" 
          CASE "NE"
               lcode(xx) ="庫房脫期"            
          CASE "D1"
 
          CASE "D0"
              if rs("amount") > 0 then
                lcode(xx) ="貸款清數"
              end if   
          case "ME","NE"
              mx = 0  
     end select 
   
     if (RS("CODE")="F3" or rs("code")="F1" or rs("code")="F2") AND mx = 0 THEN
         LNIAMT(XX)=RS("AMOUNT")
         MX = 0 
     ELSE  
         if left(rs("code"),1)="E" OR RS("code")="D0" or  rs("code")="0D" or  rs("code")="ME" or rs("code")="NE" then
            if rs("amount") <> 0 then 
               lnramt(xx) =rs("amount")
            end if         
         else
            if rs("code")="D1"  then            
               set rs1=conn.execute("select chequeamt,lnflag,appamt from loanrec where lnnum='"&rs("lnnum")&"'  ")
               if  not rs1.eof then
                   
                   lnflag = rs1(1)
                   if lnflag = "Y" then                        
                      
                      if zero = 0  then  
                         lcode(xx) ="+ 新貸  ="    
                          lniamt(xx) =  rs1(0)                                                                  
                         lnbal(xx) = lnramt(xx-1)+ lniamt(xx) 
                        
                      else
                         zero = 0
                         lniamt(xx) =  rs1(2) 
                         lcode(xx) ="+ 新貸  ="
                         lnbal(xx) = rs1(0)  
                      end if
                      bal = lnbal(xx)          
                    ELSE
                      lnbal(xx) = rs("amount")
                      lniamt(xx) =""
                      bal = lnbal(xx)
                      lcode(xx) ="新貸"  
                    END IF                     
                 end if
                rs1.close 
                MX = 0     
             end if 
           END IF
        end if
    
    xx = xx + 1
    bb = xx
    if rs("code")="D0" and rs("amount")= 0 then
       xx = xx -  1
       zero= 1
    end if
   end if 
   if left(rs("code"),1)="E"  then
      mx = mx + 1
   end if
   
   if left(rs("code"),1) ="F"  and xdate = lndate(xx-1) and mx = 1 then

         lniamt(xx-1)= rs("amount")
 
      mx = 0 
    
   end if
   if left(rs("code"),1) ="E" and mx = 2  then
      if lncode(xx-2)<> lncode(xx-1) then

      mx = mx - 1
      end if
   end if
    
   if left(rs("code"),1) ="F" and mx = 2 then
      mx = mx - 1
      lniamt(xx-1) = rs("amount")
   end if
 
 
           
   if  rs("code")="MF" or rs("code")="NF"   then
       mx = 0
       lniamt(xx-1) = rs("amount")    
   end if  

rs.movenext 
loop
rs.close
 '' response.end
end if



sql1 = "select  *  from share  where memno= '"&id&"'   order by memno,ldate,code  "
Set rs1 = Server.CreateObject("ADODB.Recordset")
rs1.open sql1, conn,2,2
yy = 1
xx = 1
xbal = 0
yy = 0
do while not rs1.eof
   xyear = year(rs1("ldate"))
   xmon  = month(rs1("ldate"))
   xday  = day(rs1("ldate"))
   xdate = xyear&xmon&xday
   ssdate = right("0"&xday,2)&"/"&right("0"&xmon,2)&"/"&xyear 
   xxdate = xyear&right("0"&xmon,2)&right("0"&xday,2)

   if xxdate  > ndate then

 	  if yy = 0 then
 	  sbal(xx) = xbal
          sdate(xx) = ydate
          samt(xx) = ""         
	   xx = xx + 1	
	   yy = 1
	 end if
     select case rs1("code")
        case "G0" ,"H0","B0","B1","B3","BE","BF","G3","H3" 
               xbal=xbal-rs1("amount")
        case   "AI"
        case "0A","A1","A2","A3","C0","C1","C3" ,"B6"  
                xbal = xbal + rs1("amount")
        end select
        sbal(xx) = xbal
        select case rs1("code")

           case "A1"
               scode(xx) = "銀行轉帳"
          case "A2"
		scode(xx) ="庫房轉帳"
          case "A3"
		scode(xx) ="現金存款"

          case "B0"
               scode(xx)="股金還款"
          case "B1"
               if rs1("amount") >0  then  
                   scode(xx)="退股"
                else
                    scode(xx)="退還貸款"
               end if
          case "B1"
                 scode(xx)="退還貸款"
          CASE "AI"
                scode(xx) ="脫期" 
          CASE "D1"
               scode(xx) ="新貸"  
          CASE "B0"
                scode(xx) ="現金退股"
         case "B3"
                scode(xx) ="退還現金"
                
          case "C0"
               scode(xx)="股息"
         
          case "C1"
               scode(xx)="股息銀行支帳" 
          case "C3"
             scode(xx)="股息現金支帳" 

          case "G0","G1","G2","G3"
               scode(xx) = "協會費"
          case "H0","H1","H2","H3"
             scode(xx) = "入會費" 
      
    end select
        sdate(xx) = ssdate
    
        samt(xx) =rs1("amount")

        xx = xx + 1
        aa = xx
  else
 
        if  left(rs1("code"),1)="G" or left(rs1("code"),1)="H" or rs1("code")="B0" or rs1("code")="B1"  then        
               xbal=xbal-rs1("amount")
        else       
              xbal = xbal +rs1("amount")
         end if
        
  end if
 
   
   rs1.movenext
loop
rs1.close
xx = 1
dim guid(3),guname(3),gusave(3),guoid(3),guoname(3),guosave(3)
xx = 1

if aa > bb then
   smax = aa
else
   smax = bb
end if
  
%>
<html>
<head>
<title>社員資料查詢</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
<center>
<table border="0" cellpadding="0" cellspacing="0">
	<tr height="30" valign="top" align="center">
		<td colspan="15"><font size="4">水務署員工儲蓄互助社</font></td>		
                <td width="200" align="right">日期 : <%=mndate%></td>  
	</tr>
	<tr height="35" valign="top" align="center">
		<td colspan="15"><font size="4">社員資料查詢</font></td>
	</tr>
	<tr height="35" valign="top" align="center">
		<td colspan="15"><font size="4">社員名稱<%=memname%><%=memcname%> </font></td>
	</tr>
	<tr height="35" valign="top" align="center">
		<td colspan="15"><font size="4">社員編號<%=id%></font></td>
	</tr>
</table>


<table border="0" cellpadding="0" cellspacing="0">
<%if guid1 <>"" then %>
        <tr>
             <td width="70"><b>1.擔保人</b><td>
             <td width="50"><%=guid1%></td>
             <td width="200"><%=guname1%></td>
             <td width="100"><b>儲蓄結餘</b></td>
             <td width="150"  align="right" ><%=formatnumber(gusave1,2)%></td>
      
        </tr>
<%end if%>
<%if guid2 <>"" then %>
        <tr>
             <td width="70"><b>2.擔保人 </b><td>
             <td width="50"><%=guid2%></td>
             <td width="200"><%=guname2%></td>
             <td width="100"><b>儲蓄結餘</b></td>
             <td width="150" align="right" ><%=formatnumber(gusave2,2)%></td>
      
        </tr>
<%end if%>
<%if guid3 <>"" then %>
        <tr>
             <td width="70"><b>3.擔保人 </b><td>
             <td width="50"><%=guid3%></td>
             <td width="200"><%=guname3%></td>
             <td width="100"><b>儲蓄結餘</b></td>
             <td width="150" align="right" ><%=formatnumber(gusave3,2)%></td>
      
        </tr>
<%end if%>
<%if guoid1 <>"" then %>
        <tr>
             <td width="90"><b>1.擔保其他人 </b><td>
             <td width="50"><%=guoid1%></td>
             <td width="300"><%=guoname1%></td>
             <td width="100"><b>貸款編號</b></td>
             <td width="50"><%=guln%></td>
      
        </tr>
<%end if%>
<%if guoid2 <>"" then %>
        <tr>
             <td width="90"><b>2.擔保其他人 </b><td>
             <td width="50"><%=guoid2%></td>
             <td width="300"><%=guoname2%></td>
             <td width="100"><b>貸款編號</b></td>
             <td width="50"><%=guln2%></td>
      
        </tr>
<%end if%>
<%if guoid3 <>"" then %>
        <tr>
             <td width="90"><b>3.擔保其他人 </b><td>
             <td width="50"><%=guoid3%></td>
             <td width="300"><%=guoname3%></td>
             <td width="100"><b>貸款編號</b></td>
             <td width="50"><%=guln3%></td>
      
        </tr>
<%end if%>
</table>
</center>

<table border="0" cellspacing="1" cellpadding="4" align="center" >
	<tr bgcolor="#330000" align="center">
		
		<td><font size="3" color="#FFFFFF">日期</font></td>
               
		<td><font size="3" color="#FFFFFF">股金</font></td>
               
		<td><font size="3" color="#FFFFFF">結餘</font></td>
                            
		<td><font size="3" color="#FFFFFF">類別</font></td>
               
                <td bgcolor="#330000"> </td>	
		<td><font size="3" color="#FFFFFF">日期</font></td>
               
		<td><font size="3" color="#FFFFFF">貸款編號</font></td>
                
		<td><font size="3" color="#FFFFFF">利息</font></td>
               
		<td><font size="3" color="#FFFFFF">每月還款<font></td>
                
		<td><font size="3" color="#FFFFFF">新貸總額/結餘</font></td>
              
		<td><font size="3" color="#FFFFFF">類別</font></td>            
	</tr>
	
	
<%
xx = 1
do while xx <= smax

if sdate(xx)<>"" or lnnum(xx)<>"" then
%>
<tr bgcolor="#FFFFF">
<%if sbal(xx) > 0   then %>
  		<td><font size="2"><%=sdate(xx)%></font></td>
               
                <%if samt(xx) <> ""   then %>
 		<td align="right"><font size="2"><%=formatNumber(samt(xx),2)%></font></td>
                <%else%>
                <td></td>
                <%end if %>  

                             
                <%if sbal(xx) <> ""   then %>
 		<td align="right"><font size="2"><%=formatNumber(sbal(xx),2)%></font></td>
                <%else%>
                <td></td>
                <%end if %>     


 		
 		<td><font size="2"><center><%=scode(xx)%></center></font></td>
<%else%>
            <td></td>
            <td></td>
            <td></td>     
             <td></td>                                    
<%end if%> 
                <td bgcolor="red"> </td>	

                <%if lnnum(xx)<>"" then %>                 
                   
		<td><font size="2"><%=Lndate(xx)%></font></td>	
               	
		<td align="right"><font size="2"><%=lnnum(xx)%></font></td>
               	
                <%if lniamt(xx)<>"" then %>
		<td align="right"><font size="2"><%=formatNumber(lniamt(xx),2)%></font></td>
                <%else%>
		<td align="right"><%=lniamt(xx)%></td>
                <%end if%>
	
                <%if lnramt(xx) <> ""   then %>           
 		     <td align="right"><font size="2"><%=formatNumber(lnramt(xx),2)%></font></td>
<%
                else
                if lcode(xx)="+ 新貸  =" then
%>                      
                   <td ><%=lcode(xx)%></td>
<% 
                   lcode(xx)=""
                   else
%>
                   <td></td>
                
<%                end if
                  end if
%>                                 
               	
		<%if lnbal(xx)<>"" then %>                
		<td align="right"><font size="2"><%=formatNumber(lnbal(xx),2)%></font></td>
                <%else%>
		<td><%=lnbal(xx)%></td>
                <%end if%>
		
		<td><font size="2"><center><%=lcode(xx)%></center></font></td>
               <%end if%>
                               

	</tr>
<%
else
  xx = 501
end if
 xx = xx + 1

 
loop
%>

</table>

</center>
</body>
</html>
<%

set rs=nothing
conn.close
set conn=nothing
%>
