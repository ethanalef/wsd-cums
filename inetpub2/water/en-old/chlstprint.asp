<!-- #include file="../conn.asp" -->

<%
id  =request.form("memno")



chkdate = request.form("stdate")


set rs=conn.execute("select a.*,b.memcname,b.memGender from crash a,memmaster b where a.memno='"&id&"' and a.memno=b.memno  ")

if  rs.eof then
    
    response.redirect "chlst.asp"
else
			For Each Field in rs.fields
			if Field.name="csdate"  then
					TheString = "if rs(""" & Field.name & """)<>"""" then " & Field.name & " = right(""0""&day(rs(""" & Field.name & """)),2)&""/""&right(""0""&month(rs(""" & Field.name & """)),2)&""/""&year(rs(""" & Field.name & """)) end if"
				else
					TheString = Field.name & "= rs(""" & Field.name & """)"
				end if
				Execute(TheString)
			Next
		
  
   if memGender="M" then
    name = memcname&"����"
   else
    name=memcname&"�k�h"
   end if
   ttlamt = lnamt+intamt
   xcsdate = rs("csdate")
   cxyr = year(xcsdae)
   cxmn = month(xcsdate)
   cxdd = day(xcsdate) 
 
end if
rs.close


mndate=right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
xmon = mid(chkdate,4,2)
xyr1  = right(chkdate,4) 
dd    = left(chkdate,2)

xdd    = dd - 1
xlnum = ""
mdate =xyr1&right("0"&xmon,2)&"01"
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
if xmdatte<= "2008/04"  then
   ndate ="20080430"
else
ndate =xyr1&right("0"&xmn,2)&right("0"&xdd,2)
end if


Set rs = server.createobject("ADODB.Recordset")
sql = "select memno,memname,memcname from memmaster where memno='"&id&"' "
rs.open sql,conn,1,1
if not rs.eof then
   memname = rs("memname")
   memcname= rs("memcname")

end if
rs.close

if memname="" then
    response.redirect "acdetaillst.asp"
   
end if
xlnnum = ""
sql = "select lnnum,repaystat from loanrec where repaystat='N' and  memno = "& id
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
dim xcode(500)
dim ldate(500)


scode(1) = "�Ѫ����l"


if xlnnum <> "" then
SQl = "select  *,convert(char(10),ldate,102) as Expr1  from loan  where memno='"&id&"' and ldate  >='"&xldate&"' and ldate<= '"&xcsdate&"' order by memno,ldate,right(code,1),left(code,1) "
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
           
          case "0D"
   		lnbal(xx)=0
        	bal = rs("amount")
                lnbal(xx)=""
          case  "D9" 
                 set ms=conn.execute("select * from loanrec where lnnum='"&rs("lnnum")&"'  ")
                 bal = ms("appamt")
                 ms.close
          case  "E0","E1" , "E4" , "E2" , "E3" , "E6" ,"E7"
                bal = bal - rs("amount")
                lnbal(xx) = bal
          case  "ER"
               bal = bal + rs("amount")
                lnbal(xx) = bal                
          case  "DE" 
                lnbal(xx) = bal
       
   end select 

   if rs("code")="F3" and  xdate <> lndate(xx-1) then

      mx = 0
   end if

   if left(rs("code"),1) ="E" or rs("code")="0D" or rs("code")="D0" or rs("code")="D1" OR rs("code")="DE" or rs("code")="NE" or  ((rs("code")="F3"  or rs("code")="F1" or rs("code")="F2" ) and MX=0)  then
      lnnum(xx) = rs("lnnum")
      newln = 0
      xyear = year(rs("ldate"))
      xmon  = month(rs("ldate"))
      xday  = day(rs("ldate"))
      oldamt = 0
      xcode(xx) = rs("code")
      ldate(xx) = rs("ldate")
      lndate(xx) =right("0"&xday,2)&"/"&right("0"&xmon,2)&"/"&xyear
      xdate      =right("0"&xday,2)&"/"&right("0"&xmon,2)&"/"&xyear     
      select case rs("code")
          case "0D"
               lcode(xx) ="�U�ڵ��l"
          case "E1"
               lcode(xx) = "�Ȧ���b"
          case "E2"
		 lcode(xx) ="�w����b"
          case "E3"
		 lcode(xx) ="�{���ٴ�"
          case "E0"
               if rs("amount") > 0 then
  		   lcode(xx) ="�Ѫ��ٴ�"
               else
                  lcode(xx) ="�h��"
               end if 
                  
          case "E6"
                    lcode(xx) ="�h��"
          case "E7"
                  lcode(xx) ="�վ�"

          case "F1"
                 lcode(xx) ="�Ȧ��ٮ�"
          case "F2"
                 lcode(xx) ="�w���ٮ�"
          case "F3"
                 lcode(xx) ="�{���ٮ�"
          case "ER"
		 lcode(xx) ="�h�٥���"
          case "F3"
                 lcode(xx) ="�{���ٮ�"  
          case "FR"
		 lcode(xx) ="�h�٧Q��"
          CASE "DE"
               lcode(xx) ="�Ȧ���" 
          CASE "DF"
               lcode(xx) ="�Q�����" 
          CASE "NE"
               lcode(xx) ="�w�в��"            
          CASE "D1"
 
          CASE "D8"
              if rs("amount") > 0 then
                lcode(xx) ="�U�ڲM��"
              end if   
          case "DE","NE"
              mx = 0  
     end select 
   
     if (RS("CODE")="F3" or rs("code")="F1" or rs("code")="F2") AND mx = 0 THEN
         LNIAMT(XX)=RS("AMOUNT")
         MX = 0 
     ELSE  
         if left(rs("code"),1)="E" OR RS("code")="D8" or  rs("code")="0D" or  rs("code")="DE" or rs("code")="NE" then
       
            if rs("amount") <> 0 then 
               if RS("code")="0D" then
                   lnbal(xx) = rs("amount")
               else
                  lnramt(xx) =rs("amount")
               end if
            end if         
         else
            if rs("code")="D1"  then            
               set rs1=conn.execute("select chequeamt,lnflag,appamt,loantype from loanrec where lnnum='"&rs("lnnum")&"'  ")
               if  not rs1.eof then
                   loantype = rs1("loantype")   
                             
                   lnflag = rs1(1)
                   if lnflag = "Y" then                        
                      
                      if rs1(0) <> 0 then  
                         lniamt(xx) =  rs1(0) 
                      else
                         lniamt(xx) =""
                      end if
                     
                         lcode(xx) ="+ �s�U  ="
                      
                         lnbal(xx) = rs1(2)  
                    
                      bal = lnbal(xx)          
                    ELSE
                      lnbal(xx) = rs1("appamt")
                      lniamt(xx) =""
                      bal = lnbal(xx)
                      lcode(xx) ="�s�U"  
                    END IF                     
                 end if
                rs1.close 
                MX = 0     
             end if 
           END IF
        end if
    
    xx = xx + 1
    bb = xx
    if rs("code")="D8" and rs("amount")= 0 then
       xx = xx -  1
       zero= 1
    end if
   end if 
   if  rs("code")="DF" or rs("code")="NF"   then
       mx = 0
       lniamt(xx-1) = rs("amount")    
   end if  

  if left(rs("code"),1) ="E"   then
     mx = mx + 1
     tmpdate =  rs("ldate") 
  
     rs.movenext
   
     if not rs.eof then    
        select case   left(rs("code"),1)
            case "E"
                 if rs("ldate") <> tmpdate then
                    mx = 0
                                 
                     
                 end if
                 rs.moveprevious  
                   
                 
               
            case "F"
                              
                     if mx = 1 then
                         lniamt(xx-1)= rs("amount") 
                         mx = 0
                     else
                       select case rs("code")
                              case "F0","F3"
                                    lniamt(xx-2)= rs("amount")
                                    if  not rs.eof then
                                    rs.movenext
                                    if left(rs("code"),1)="F" then
                                       lniamt(xx-1)= rs("amount")
                                    else
                                       rs.moveprevious  
                                    end if 
                                    end if
                              case "F1"
                                    lniamt(xx-2)= rs("amount")
                                    if  not rs.eof then
                                    rs.movenext
                                     if left(rs("code"),1)="F" then
                                        lniamt(xx-1)= rs("amount")
                                     else
                                        rs.moveprevious  
                                    end if  
                                    end if                 
                              case "F2"
                                   lniamt(xx-1)= rs("amount") 
                                                  
                        end select             
                        mx = 0

                    end if
                         
         case else
                mx = 0
                rs.moveprevious  

         end select 
         
    end if
  end if

  

          

if  not rs.eof then
rs.movenext 
end if
loop
rs.close
 '' response.end
end if




sql1 = "select  * from share  where memno= '"&id&"'  and ldate<= '"&xcsdate&"'  order by memno,ldate,code  "
Set rs1 = Server.CreateObject("ADODB.Recordset")
rs1.open sql1, conn,2,2

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


   if ( xxdate  > ndate) then

 	  if sbal(xx) > 0 and  yy = 0 then
   
  	    xx = xx + 1	
            yy = 1
	 end if
    select case rs1("code")
        case "G0" ,"H0","B0","B1","B3","BE","BF","G3","H3" 
               sbal(xx)=sbal(xx-1)-rs1("amount")
        case   "AI","CH"
                   sbal(xx)=sbal(xx-1)
        case  "A1","A2","A3","C0","C1","C3" ,"A0","A7" ,"A4" 
                sbal(xx) = sbal(xx-1) + rs1("amount")
        end select
      
        select case rs1("code")
          
           case "A0"
               scode(xx) = "�h�ٶU��"
           case "A1"
               scode(xx) = "�Ȧ���b"
          case "A2"
		scode(xx) ="�w����b"
          case "A3"
		scode(xx) ="�{���s��"
          case "A4"
              scode(xx) ="�O�I��"
          case "B0"
               scode(xx)="�Ѫ��ٴ�"
          case "A7"
                  scode(xx) ="�վ�"
          case "B1"
             
                   scode(xx)="�h��"

          CASE "AI"
                scode(xx) ="����@�@" 
          CASE "D1"
               scode(xx) ="�s�U�Ȧ�"  
          CASE "B0"
                scode(xx) ="�{���h��"
         case "B3"
                scode(xx) ="�h�ٲ{��"
                
          case "C0"
               scode(xx)="�Ѯ��@�@"
          case "CH"
               scode(xx)="�Ȱ��Ѯ�"            
          case "C1"
               scode(xx)="�Ѯ��Ȧ�" 
          case "C3"
             scode(xx)="�Ѯ��{��" 

          case "G0 ","G1","G2","G3"
                scode(xx) = "�J���O"
          case "H0","H1","H2","H3"
            
              scode(xx) = "��|�O" 
    end select
        sdate(xx) = ssdate
    if rs1("amount") <> "" then       
        samt(xx) =rs1("amount")
    end if
        xx = xx + 1
        aa = xx
  else

        if  left(rs1("code"),1)="G" or left(rs1("code"),1)="H" or left(rs1("code"),1)="B" then        
               sbal(xx)=sbal(xx)-rs1("amount")
        else 
        if rs1("code")="0A" then          
            sbal(xx) = sbal(xx) +rs1("amount")
        end if  
	
            sdate(xx) = ssdate
            scode(xx) = "�Ѫ����l" 
            samt(xx)  =  ""
             aa = 1
         end if
        
  end if
 
   
   rs1.movenext
loop
rs1.close
xx = 1
if aa > bb then
   smax = aa
else
   smax = bb
end if
if sbal(1) = 0 then
   sbal(1) =""
   scode(1)=""
   
end if
if request.form("output")="word" then
	Response.ContentType = "application/msword"
elseif request.form("output")="excel" then
	Response.ContentType = "application/vnd.ms-excel"
end if
  
%>
<html>
<head>
<title>������Ƭd��</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
<center>
<br>
<BR>
<table border="0" cellpadding="0" cellspacing="0">
	<tr height="30" valign="top" align="center"><td colspan="15"><font size="4"  face="�з���" >���ȸp���u�x�W���U��<br>������Ƭd��<br><font size="2"  face="�з���" >��� : <%=mndate%></font></font></td></tr>
        <tr height="30" valign="top" align="center"><td colspan="15"><font size="4"  face="�з���" >�����W�� : <%=memname%> (<%=memcname%>)  �����s�� : <%=id%></td> 
        <tr height="20" ><td colspan=9></td></tr>

        


<%if guid1 <>"" then %>
        <tr>
             <td width="70"><b>1.��O�H</b><td>
             <td width="50"><%=guid1%></td>
             <td width="200"><%=guname1%></td>
             <td width="100"><b>�x�W���l</b></td>
             <td width="150"  align="right" ><%=formatnumber(gusave1,2)%></td>
      
        </tr>
<%end if%>
<%if guid2 <>"" then %>
        <tr>
             <td width="70"><b>2.��O�H </b><td>
             <td width="50"><%=guid2%></td>
             <td width="200"><%=guname2%></td>
             <td width="100"><b>�x�W���l</b></td>
             <td width="150" align="right" ><%=formatnumber(gusave2,2)%></td>
      
        </tr>
<%end if%>
<%if guid3 <>"" then %>
        <tr>
             <td width="70"><b>3.��O�H </b><td>
             <td width="50"><%=guid3%></td>
             <td width="200"><%=guname3%></td>
             <td width="100"><b>�x�W���l</b></td>
             <td width="150" align="right" ><%=formatnumber(gusave3,2)%></td>
      
        </tr>
<%end if%>
<%if guoid1 <>"" then %>
        <tr>
             <td width="90"><b>1.��O��L�H </b><td>
             <td width="50"><%=guoid1%></td>
             <td width="300"><%=guoname1%></td>
             <td width="100"><b>�U�ڽs��</b></td>
             <td width="50"><%=guln%></td>
      
        </tr>
<%end if%>
<%if guoid2 <>"" then %>
        <tr>
             <td width="90"><b>2.��O��L�H </b><td>
             <td width="50"><%=guoid2%></td>
             <td width="300"><%=guoname2%></td>
             <td width="100"><b>�U�ڽs��</b></td>
             <td width="50"><%=guln2%></td>
      
        </tr>
<%end if%>
<%if guoid3 <>"" then %>
        <tr>
             <td width="90"><b>3.��O��L�H </b><td>
             <td width="50"><%=guoid3%></td>
             <td width="300"><%=guoname3%></td>
             <td width="100"><b>�U�ڽs��</b></td>
             <td width="50"><%=guln3%></td>
      
        </tr>
<%end if%>
</table>
</center>

<table border="0" cellspacing="1" cellpadding="1" align="center" >
	<tr >
		
		<td width=80 align="center">���</td>
               
		<td width=60 align="center">�Ѫ�</td>
               
		<td width=100 align="center">���l</td>
                            
		<td width=80 align="center">���O</td>
               
                <td  width=1 align="center"> </td>	
		<td width=80 align="center">���</td>
               
		<td width=80 align="center">�ɾڽs��</td>
                
		<td width=60 align="center">�Q��</td>
               
		<td width=60 align="center">�C���ٴ�</td>
                
		<td width=100 align="center">�s�U�`�B/���l</td>
              
		<td width=100 align="center">���O</td>            
	</tr>
	
	<tr><td colspan=11><hr></td></tr>	
<%
xx = 1
do while xx <= smax

if sdate(xx)<>"" or lnnum(xx)<>"" then
%>
<tr bgcolor="#FFFFF">
<%if sbal(xx) <> ""   then %>
  		<td width=80 align="center"><%=sdate(xx)%></td>
               
                <%if samt(xx) <> ""   then %>
 		<td width=60 align="right"><%=formatNumber(samt(xx),2)%></td>
                <%else%>
                <td></td>
                <%end if %>  

                             
                <%if sbal(xx) <> ""   then %>
 		<td width=80 align="center"><%=formatNumber(sbal(xx),2)%></td>
               
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
                <td colour="red">�x</td>	

                <%if lnnum(xx)<>"" then %>                 
                   
		<td width=80 align="center"><%=Lndate(xx)%></td>	
               	
		<td width=80 align="center"><%=lnnum(xx)%></td>
               	
                <%if lniamt(xx)<>"" then %>
		<td width=60 align="right"><%=formatNumber(lniamt(xx),2)%></td>
                <%else%>
		<td align="right"><%=lniamt(xx)%></td>
                <%end if%>
	
                <%if lnramt(xx) <> ""   then %>           
 		    <td width=60 align="right"><%=formatNumber(lnramt(xx),2)%></td>
<%
                else
                if lcode(xx)="+ �s�U  =" then
                   if loantype = "E" then
                      lcode(xx)=" ������ "
                   end if
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
		<td width=100 align="right"><%=formatNumber(lnbal(xx),2)%></td>
                <%else%>
		<td><%=lnbal(xx)%></td>
                <%end if%>
		
		<td width=100 align="center"><%=lcode(xx)%></font></td>
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
<br>
<br>
<font size=2>
<div>�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�ѩ󦬨����<%=name%>���}���ӽгq��(�ɮ׽s��<font size=1><b><%=refyr%></font></b>�~<font size=1><b><%=refno%></font></b>�v)�A�]����]���}�l</div>
<div>�@�@�@�@�@�@�@�@�@�@�@�@�@�ᵲ�A�ӺI��<font size=1><b><%=year(xcsdate)%></font></b>�~<font size=1><b><%=cxmn%></font></b>��<font size=1><b><%=cxdd%></font></b>���A<%=name%>�b�������Ѫ��s�ڤ���<font size=1.5><b>$<%=formatNumber(shamt,2)%></font></b>�����s�A</div>
<div>�@�@�@�@�@�@�@�@�@�@�@�@�@�ө|��U�ڵ��l����<font size=1><b>$<%=formatNumber(ttlamt,2)%></font></b>�C�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@</div>
<div>�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�����b����<%=name%>�����}���q������A���g�l�}���Ʒ|��ĳ�A�P�|�̬ҳq�L�Y��</div>
<div>�@�@�@�@�@�@�@�@�@�@�@�@�@������ӭ���k��<font size=1><b>119</font></b>����<font size=1><b>3</font></b>��<font size=1><b>22</font></b>�q</font></b>2</font></b>���Ҹ����ѥ��d�m�v�A�Y<%=name%>���Ѫ����l���ơ@�@</div>
<div>�@�@�@�@�@�@�@�@�@�@�@�@�@<font size=1><b>($<%=formatNumber(shamt,2)%>)</font></b>�Χ@��P�Ӥ�ڡA�g�ֺ�p�U�G<font size=1><b>$<%=formatNumber(shamt,2)%> - $<%=formatNumber(ttlamt,2)%> = -($<%=formatNumber((shamt-ttlamt)*-1,2)%>)</font></b>�A</div>

<div>�@�@�@�@�@�@�@�@�@�@�@�@�@�P�ɲפ����y�A�ö����ӭ���k��<font size=1><b>119</font></b>����<font size=1><b>3</font></b>��<font size=1><b>21</font></b>�q�Ҹ��A<%=name%>�������M�t�����y��</div>
<div>�@�@�@�@�@�@�@�@�@�@�@�@�@��ɪ��k��(�Ű�)�d���F�]���A�H�W<font size=1><b>$<%=formatNumber((shamt-ttlamt)*-1,2)%></font></b>��<%=name%>���̫��ڵ��l�A�����N�Q��</div>
<div>�@�@�@�@�@�@�@�@�@�@�@�@�@���v�ҩ���(�@��q�Ϊ��<font size=1><b>46A</font></b>)�A�t�[�Ʈ׶O<font size=1><b>$40</font></b>)�A�V�}�׺޲z�x�e���A�åѸӸp�w�Ƹ�i�Ʃy�C</div>
</font>
</font>
</center>
</body>
</html>
<%

set rs=nothing
conn.close
set conn=nothing
%>
