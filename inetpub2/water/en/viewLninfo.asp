<!-- #include file="../conn.asp" -->
<!-- #include file="init.asp" -->
<%
Function LoanCode(ByVal x)
	select case x
          case "0D"
				LoanCode="�U�ڵ��l"
          case "E1"
				LoanCode= "�Ȧ���b"
          case "E2"
				LoanCode="�w����b"
          case "EC"
                LoanCode="�������B"
          case "E3"
				LoanCode="�{���ٴ�"
          case "E0"
               if rs("amount") > 0 then
					LoanCode="�Ѫ��ٴ�"
               else
					LoanCode="�h�٥���"
               end if 
          case "F0"
              if rs("amount") > 0 then
					LoanCode="�Ѫ��ٴ�"
               else
					LoanCode="�h�٧Q��"
               end if 
          case "E6"
                LoanCode="�h��"
          case "E7"
				LoanCode="�վ�"
          case "F1"
				LoanCode="�Ȧ��ٮ�"
          case "F2"
				LoanCode="�w���ٮ�"
          case "F3"
                LoanCode="�{���ٮ�"
          case "ER"
				LoanCode="�h�٥���"
          case "F3"
                LoanCode="�{���ٮ�"  
          case "FR"
				LoanCode="�h�٧Q��"
          CASE "DE"
				LoanCode="�Ȧ���" 
          CASE "DF"
				LoanCode="�Q�����" 
          CASE "NE"
				LoanCode="�w�в��"            
          CASE "D9"
 
          CASE "D8"
              if rs("amount") > 0 then
                LoanCode="�U�ڲM��"
              end if   
          case "DE","NE"
              mx = 0  
	end select
End Function

memno =request("key")
pos = instr(memno,"*")
chkdate=right(memno,10)
id = left(memno,pos-1)

if id="" then%>
	<script language="JavaScript">
	<!--
		window.opener.document.form1.memNo.focus()        
		window.close();
	//-->
	</script>

	<%response.end
end if

mndate=right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
xmon  	= mid(chkdate,4,2)
xyr1 	= right(chkdate,4) 
dd    	= left(chkdate,2)
xlnum 	= ""
mdate 	= dateserial(xyr1,xmon,0)
opdate 	= dateserial(2008,4,30)
xmdate=xyr1&right("0"&xmon,2)
ydate =xdd&"/"&right("0"&xmn,2)&"/"&xyr1
if mdate<= opdate  then

   ndate =dateserial(2008,04,30)
else
ndate =mdate
end if


Set rs = server.createobject("ADODB.Recordset")
sql = "select memno,memname,memcname,accode from memmaster where memno='"&id&"' "
rs.open sql,conn,1,1
if not rs.eof then
   memname = rs("memname")
   memcname= rs("memcname")
   xaccode = rs("accode")
end if
rs.close

if xaccode<>"9999" then
set rs=conn.execute("select memcname,memname,memofficetel from memmaster where memno='"&xaccode&"' ")
if not rs.eof then
   xaccname = rs("memcname")
   xacname  = rs("memname")
   xactel   = rs("memofficetel")
end if
rs.close
else
   xaccname = "�u�@�H��"
   xacname  = ""
   xactel   = "27879222"
end if

xlnnum = ""


set rs1 = conn.execute("select  lnnum,lndate  from loanrec where memno='"&id&"' and repaystat='N'  ")
if not rs1.eof then
   xldate = rs1(1)
   xlnnum = rs1(0)
end if
rs1.close 
   


if xlnnum <> "" then
   sql = " select a.* from guarantor a,loanrec b where a.lnnum=b.lnnum and b.repaystat='N' and a.lnnum = "& xlnnum
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
									case "G","H","B","M"
										 bal = bal - rs1("amount")
									case "A","T","C","0"
										 if rs1("code")<>"AI" AND CODE<>"CH" then
											bal = bal+ rs1("amount")
										 end if
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
								case "G","H","B","M"
									 bal = bal - rs1("amount")
								case "A","T","C","0"
									 if rs1("code")<>"AI" AND CODE<>"CH" then
										bal = bal+ rs1("amount")
									 end if
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
								case "G","H","B","M"
									 bal = bal - rs1("amount")
								case "A","T","C","0"
									if rs1("code")<>"AI" and code <>"CH" then
										bal = bal+ rs1("amount")
									 end if
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

sql = "select a.* from guarantor a,loanrec b  where  a.lnnum=b.lnnum and b.repaystat='N' and a.guarantorID = "&id
rs.open sql, conn, 1, 1
do while  not rs.eof 
     select case xx
             case 1 
                  guoid1 = rs("memno")                              
                  if guoid1 <> "" then
					  guln1=""
					  Set rs1 = server.createobject("ADODB.Recordset")
					  sql1 = " select * from loanrec where repaystat='N' and memno = '"& guoid1&"' "    
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


set ms=conn.execute("select lnnum from loan where memno='"&id&"' and ldate>='"&mdate&"' order by memno,ldate,code ")
if not ms.eof then
   xlnnum = ms("lnnum")
 

end if
ms.close

SQl = "select lnnum,ldate  from loan  where   lnnum='"&xlnnum&"' order by lnnum,ldate,right(code,1) desc " 
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
dim xcode(500)
dim lnramt(500)
dim lniamt(500)
dim lnbal(500)
dim lncode(500)
dim ldate(500)

scode(1) = "�Ѫ����l"


if xlnnum <> "" then
SQl = "select  *  from loan  where memno='"&id&"'  order by memno,ldate,uid,right(code,1),left(code,1) "
''response.write(sql)'
''response.end

Set rs = Server.CreateObject("ADODB.Recordset")	
rs.open sql, conn,2  

cc = 0
xx = 1
qx = 0 = 0
MX=0
zero = 0
    

  
do while not rs.eof
   xyear = year(rs("ldate"))
   xmon  = month(rs("ldate"))
   xday  = day(rs("ldate"))
   

   ssdate = right("0"&xday,2)&"/"&right("0"&xmon,2)&"/"&xyear  
  if rs("ldate") > ndate  then
       
	  if lnbal(xx) > 0 and  yy = 0 then
               bal = lnbal(xx)
 
  	    xx = xx + 1	
            yy = 1
	 end if
  lncode(xx) = rs("code")
  xdate = right("0"&day(rs("ldate")),2)&"/"&right("0"&month(rs("ldate")),2)&"/"&year(rs("ldate"))
        select case rs("code")
            case "0D"
				lnbal(xx)=rs("amount")
				bal = rs("amount")
                lnbal(xx)=bal 
            case  "E0","E1" , "E4" , "E2" , "E3" , "E6" ,"E7","EC"
                lnramt(xx) =rs("amount")
                bal = bal - rs("amount")
                lnbal(xx) = bal
            case  "ER"
               lnramt(xx) =rs("amount")
               bal = bal + rs("amount")
               lnbal(xx) = bal                
            case  "DE" 
              
				lnramt(xx) =rs("amount")
                lnbal(xx) = bal
			case  "DF" 
              
				lniamt(xx) =rs("amount")
                lnbal(xx) = bal	
				
                 
            case  "D8"
                  lnramt(xx) =rs("amount")
                  bal = 0
                  lnbal(xx) = 0        
           end select 

		   
       if left(rs("code"),1) ="E" or rs("code")="0D" or rs("code")="D8" or rs("code")="D9" OR rs("code")="DE" or rs("code")="NE" OR rs("code")="DF" then
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
              lcode(xx) = LoanCode(rs("code")) 
           xx = xx + 1          
           bb = bb + 1
       end if   
       if (rs("code")="F3"  or rs("code")="F1" or rs("code")="F2" or rs("code")="F7" or rs("code")="F0" or rs("code")="FR" )  then
 
       
           if ldate(xx-1) = rs("ldate") and xcode(xx-1) = "E"&right(rs("code"),1) then
				lniamt(xx-1) = rs("amount")
           end if
           if ldate(xx-1) <> rs("ldate") and (left(rs("code"),1) = "F" and code<>"F3" )  then
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
				  lcode(xx) = LoanCode(rs("code")) 
				  lniamt(xx) = rs("amount")
				  lnbal(xx) = lnbal(xx-1)
			   xx = xx + 1          
			   bb = bb + 1

           end if
           if rs("code") = "F3" then
				  xidate = rs("ldate")
				  set mr1 = conn.execute("select count(*) from loan where memno='"&rs("memno")&"' and ldate = '"&xidate&"'  and code='F3' group by memno,ldate,code ")
				  xcnt = mr1(0)
				  do while xcnt > 0 and not rs.eof
					 lniamt(xx - xcnt)= rs("amount")
					 xcnt = xcnt -1
					 if xcnt > 0 then
					   rs.movenext
					 end if
				   loop
	 
			   end if           
         end if
 

         if rs("code")="D9"  then            
                   set rs1=conn.execute("select chequeamt,lnflag,appamt,loantype from loanrec where lnnum='"&rs("lnnum")&"'  ")
                      if  not rs1.eof then
                          loantype = rs1("loantype")   
                             
                          lnflag = rs1(1)
                            if lnflag = "Y" then                        
                      
                                if rs1(0) <> 0 then  
                                   lniamt(xx-1) =  rs1(0)
                                else
                                   lniamt(xx-1) =""
                                end if
                                if loantype ="N" then
                        
                                   lcode(xx-1)  = "+ �s�U  ="
                                else
                     
                                   lcode(xx-1)=" ������ "
                                end if               
                      
                                lnbal(xx-1) = rs1(2)  
                    
                                bal = lnbal(xx-1)      
   
                             ELSE
                                 lnbal(xx-1) = rs1("appamt")
                                 lniamt(xx-1) =""
                                 bal = lnbal(xx-1)
                                 lcode(xx-1) ="�s�U"  
                            END IF                     
                         end if
                         rs1.close 
                         MX = 0     
   
                      end if 

        
    
 
'  if  rs("code")="DF" or rs("code")="NF"   then
      ' mx = 0
 '     lniamt(xx-1) = rs("amount")    
  '     if rs("memno") ="4480"  and rs("code")="DF"  and  ymd(rs("ldate")) >="2016/07/30" and rs("lnnum")= "2013080003" then
   '                   lnbal(xx-1) = bal + rs("amount") 
    '                  bal =    lnbal(xx-1)
     '  end if   
  
 	'	if rs("memno") ="2527"  and rs("code")="DF"  and  ymd(rs("ldate")) >="2021/05/29" and rs("lnnum")= "2020080005" then
	'					lniamt(xx) =rs("amount")
	'					 lnramt(xx) =""
	'	end if
               
    'end if



   else

        if  left(rs("code"),1)="E" or rs("code")="D8"  then        
               lnbal(xx)=lnbal(xx)-rs("amount")
        else 
            if rs("code")="0D" or rs("code")="D9"   then 
                if rs("code") = "D9" then
                  set rs1=conn.execute("select chequeamt,lnflag,appamt,loantype from loanrec where lnnum='"&rs("lnnum")&"'  ")
                      if  not rs1.eof then
                          lnbal(xx) = rs1(2)                      
                      end if
                else                                                 
                     lnbal(xx) = rs("amount")
                end if
           end if
        end if  

	    lnnum(xx) = rs("lnnum")
            lndate(xx) = ssdate
            lcode(xx) = "�U�ڵ��l" 
            lniamt(xx)  =""     
            aa = 1
     end if


if  not rs.eof then
rs.movenext 

end if
loop
rs.close
 '' response.end
end if


sql1 = "select  *  from share  where memno= '"&id&"'   order by memno,ldate,code  "
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
 

	if rs1("ldate") > ndate  then
		if sbal(xx) > 0 and  yy = 0 then
			xx = xx + 1	
            yy = 1
		end if
		select case rs1("code")
		
			case "H7","B8"
				sbal(xx)=sbal(xx-1)-rs1("amount")	
			case "G0" ,"H0","B0","B1","B3","BE","BF","G3","H3","MF" 
				   sbal(xx)=sbal(xx-1)-rs1("amount")
			case   "AI","CH"
					  sbal(xx) = sbal(xx-1)     
			case  "A1","A2","A3","C0","C1","C3" ,"A0","A7" ,"A4" ,"C5"
					sbal(xx) = sbal(xx-1) + rs1("amount")
			case "E7","A8"
					sbal(xx) = sbal(xx-1) + rs1("amount")
        end select
        select case rs1("code")
			case "A0"
					scode(xx) 	= "�h�ٶU��"
			case "A1"
					scode(xx) 	= "�Ȧ���b"
			case "A2"
					scode(xx) 	="�w����b"
			case "A3"
					scode(xx) 	="�{���s��"
			case "A4"
					scode(xx) 	="�O�I��"
			case "B0"
					scode(xx)	="�Ѫ��ٴ�"
			case "A7"
					scode(xx) 	="�վ�"
			case "B1"
					scode(xx)	="�h��"
			CASE "AI"
					scode(xx) 	="����@�@" 
			CASE "D9"
					scode(xx) 	="�s�U�Ȧ�"  
			CASE "B0"
					scode(xx) 	="�{���h��"
			case "B3"
					scode(xx) 	="�h�ٲ{��"
			case "C0"
					scode(xx)	="�Ѯ��@�@"
			case "CH"
					scode(xx)	="�Ȱ��Ѯ�"        
			case "C1"
					scode(xx)	="�Ѯ��Ȧ�" 
			case "C3"
					scode(xx)	="�Ѯ��{��" 
			case "C5"
					scode(xx)	="�Ѯ��ٴ�"
			case "G0","G1","G2","G3"
					scode(xx) 	= "�J���O"
			case "H0","H1","H2","H3"
					scode(xx) 	= "��|�O" 
			case "MF"
					scode(xx) 	= "�N��O" 
			case "E7","A8"	
					scode(xx)  ="��(+)��"
			case "H7","B8"
					scode(xx)  ="��(-)��"
	
		end select
        sdate(xx) = ssdate
		if rs1("amount") <> "" then       
			samt(xx) =rs1("amount")
		end if
        xx = xx + 1
        aa = xx
	else

        if  rs1("code")="H7" or left(rs1("code"),1)="G" or left(rs1("code"),1)="H" or left(rs1("code"),1)="B"  or  rs1("code")="MF"  then        
				sbal(xx)=sbal(xx)-rs1("amount")
        else 
			if rs1("code")="E7" or rs1("code")="0A" or left(rs1("code"),1)="C" or left(rs1("code"),1)="A"  then 
			   if rs1("code")<>"AI"  and rs1("code")<>"CH"  then
					sbal(xx) = sbal(xx) +rs1("amount")
			   end if
			end if  
			sdate(xx) = ssdate
            scode(xx) = "�Ѫ����l" 
            samt(xx)  =""     
            aa = 1
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
if sbal(1) = 0 then
   sbal(1) =""
   scode(1)=""
   
end if

%>
<html>
<head>
<title>������Ƭd��</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">

<input type="hidden" name="loantype" value="<%=loantype%>">
<center>
<font size="4"  face="�з���" >
���ȸp���u�x�W���U��
<br>
�ӤH��f�d�ߦC��

<br>

</font>
<font size="3"  face="�з���" >
��� : <%=mndate%>
</font>
<br>
<br>
<font size="4"  face="�з���" >
�����W�� : <%=memname%> (<%=memcname%>)  �����s�� : <%=id%> 
</font>
<br>
<br>                   


<table border="0" cellpadding="0" cellspacing="0">
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
             <td width="50"><%=guln1%></td>
      
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
<br>
<table border="0" cellspacing="1" cellpadding="1"   align="center" >
       <tr>
       
            <td  vlign="top"><font size="3"  face="�з���" >�p���H�G<%=xaccname%>�@<%=xacname%>�@�q�ܡG<%=xactel%></font></td>
       </tr>
</table>
<br>
<table border="0" cellspacing="1" cellpadding="1" align="center" >
	<tr >
		
		<td width=80 align="center">���</td>
               
		<td width=80 align="center">�Ѫ�</td>
               
		<td width=100 align="center">���l</td>
                            
		<td width=100 align="center">���O</td>
               
                <td  width=1 align="center"> </td>	
		<td width=80 align="center">���</td>
               
		<td width=80 align="center">�ɾڽs��</td>
                
		<td width=100 align="center">�Q��</td>
               
		<td width=100 align="center">�C���ٴ�</td>
                
		<td width=100 align="center">�s�U�`�B/���l</td>
              
		<td width=100 align="center">���O</td>            
	</tr>
	
	<tr><td colspan=11><hr></td></tr>	
<%



xx = 1
do while xx <= smax+1

if sdate(xx)<>"" or lnnum(xx)<>"" then
%>
<tr bgcolor="#FFFFF">
<%if sbal(xx) <> ""     then %>
  		<td><font size="2"><%=sdate(xx)%></font></td>
               
                <%if samt(xx) <> ""   then %>
 		<td width=80 align="right"><%=formatNumber(samt(xx),2)%></td>
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
                <td bgcolor="red"> </td>	

<%if lnnum(xx) <>""    then %>                 
        
		<td><font size="2"><%=Lndate(xx)%></font></td>	
               	
		<td align="right"><font size="2"><%=lnnum(xx)%></font></td>
               	
                <%if lniamt(xx)<>"" then %>
		<td width=100 align="right"><%=formatNumber(lniamt(xx),2)%></td>
                <%else%>
		<td align="right"><%=lniamt(xx)%></td>
                <%end if%>
	
                <%if lnramt(xx) <> ""   then %>           
 		    <td width=100 align="right"><%=formatNumber(lnramt(xx),2)%></td>
<%
                else
                if lcode(xx)="+ �s�U  =" then
 
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
            
                               

	</tr>
<%
end if
else
  xx = 501
end if
 xx = xx + 1

 
loop
%>

</table>
</font>
</center>
</body>
</html>
<%

set rs=nothing
conn.close
set conn=nothing
%>
