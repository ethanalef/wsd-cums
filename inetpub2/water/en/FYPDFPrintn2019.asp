<!-- #include file="../conn.asp" -->
<!-- #include file="init.asp" -->
<%
Function ShareCode(ByVal x)
	select case x
			case "0A"
				ShareCode = "�Ѫ����l"
			case "A0"
				ShareCode = "�h�ٶU��"
			case "A1"
				ShareCode = "�Ȧ���b"
			case "A2"
				ShareCode ="�w����b"
			case "A3"
				ShareCode ="�{���s��"
			case "A4"
				ShareCode ="�O�I��"
			case "B0"
				ShareCode="�Ѫ��ٴ�"
			case "A7"
					ShareCode ="�վ�"
			case "B1"
					ShareCode="�h��"
			CASE "AI"
					ShareCode ="����@�@" 
			CASE "D9"
					ShareCode ="�s�U�Ȧ�"  
			CASE "B0"
					ShareCode ="�{���h��"
			case "B3"
					ShareCode ="�h�ٲ{��"
			case "C0"
					ShareCode="�Ѯ��@�@"
			case "CH"
					ShareCode="�Ȱ��Ѯ�"        
			case "C1"
					ShareCode="�Ѯ��Ȧ�" 
			case "C3"
					ShareCode="�Ѯ��{��" 
 			case "G0","G1","G2","G3"
					ShareCode = "�J���O"
			case "H0","H1","H2","H3"
					ShareCode = "��|�O" 
			case "MF"
					ShareCode = "�N��O" 
	end select
End Function

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
               if ms("amount") > 0 then
					LoanCode="�Ѫ��ٴ�"
               else
					LoanCode="�h�٥���"
               end if 
          case "F0"
              if ms("amount") > 0 then
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
          CASE "D8"
                LoanCode="�U�ڲM��"
              
          case "DE","NE"
              mx = 0  
	end select
End Function

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

On Error Resume Next

server.scripttimeout = 1800
mon = mid( request.form("paidday"),4,2)
nday = left( request.form("paidday"),2)
xname = request.form("accode")
rate  = request.form("nrate")
pos = instr(xname,"-")
if pos>0 then
	memno = mid(xname,1,4)
	mname =  mid(xname,pos+1,50)
else
	response.redirect "FYPDFPrintn88.asp"
end if
if memno<>"9999"  then
	memno = left(xname,pos-1)
	mname =  mid(xname,pos+1,50)
	styfield = "accode='"&memno&"' and "
else
	styfield = "accode='9999' and "
	if instr(mname,"(1-500)") > 0 then
        styfield = styfield & "a.memno<='500' and "
	else
		if instr(mname,"(501-1000)") > 0 then
			styfield = styfield & " (a.memno>='501' and  a.memno<='1000') and "
		else 
			if instr(mname,"(1001-1500)") > 0 then
				styfield =styfield &  " (a.memno>='1001' and  a.memno<='1500') and "
			else 
				if instr(mname,"(1501-2000)") > 0 then
					styfield =styfield &  " (a.memno>='1501' and  a.memno<='2000') and "
				else 
					if instr(mname,"(2001-2500)") > 0 then
						styfield =styfield &  " (a.memno >='2001' and   a.memno<='2500') and "
					else
						if instr(mname,"(2501-3000)") > 0 then
							styfield = styfield & " (a.memno >='2501' and   a.memno<='3000') and "
						else
							if instr(mname,"(3001-5000)") > 0 then
								styfield =styfield &  " (a.memno >='3001' and   a.memno<='5000') and "
							else
								styfield = ""
							end if
						end if
					end if
				end if
			end if
		end if
	end if
end if

yy =request.form("Nyear")
pyear = yy
Nyear = (yy-1)&"/"&yy

xxdate =dateserial(yy-1,7,1)
yydate =dateserial(yy,7,1)
mmdate = (yy-1)&"/07/01"
nndate = yy&"/07/01"
chkdate ="01/07/"&(yy-1)
ndate=dateserial(2008,4,30)
if xxdate < ndate then
   xxdate = ndate
   chkdate ="30/04/"&(yy-1)
end if

if memno ="9999" then
	xmemname = ""
	xmemcname = "�u�@�H��"
   xmemContactTel = "27879222"
else
	set rs=conn.execute("select memno,memname,memcname,memofficeTel from memmaster where memno='"&memno&"' ")
if not rs.eof then
   xmemname = rs("memname")
   xmemcname = rs("memcname")
   xmemContactTel = rs("memofficeTel")
end if
rs.close

end if
dim divd(12)
dim dvdate(12)
xyear = yy - 1
for i = 1 to 12
    dvdate(i)=dateserial(xyear,6+i,1)
next  

SQl = "select memno,memname,memcname,B1, B2, B1relation, B2relation,mstatus  from memmaster where  "&styfield&"  mstatus not in ('C','P','B' ) order by memno   "
SQl = "select a.memno,a.memname,a.memcname,a.B1, a.B2, a.B1relation, a.B2relation, b.dttl ,b.cttl , c.ndttl ,c.ncttl    from memmaster a "&_
      " right join ( select memno ,sum( case when left(code,1)  in  ( '0','A','C' ) and code<>'AI' and code<>'CH' "&_
      "  then amount else 0 end ) as dttl , sum( case when left(code,1)  in  ( 'G','H','B' ) and code<>'MF' "&_
      "  then amount else 0 end ) as cttl from share where "&_
      " ldate < '"&nndate&"'   "&_
      " group by memno ) b on a.memno=b.memno "&_
      " right join ( select memno ,sum( case when left(code,1)  in  ( '0','A','C' ) and code<>'AI' and code<>'CH' "&_
      "  then amount else 0 end ) as ndttl , sum( case when left(code,1)  in  ( 'G','H','B' ) and code<>'MF' "&_
      "  then amount else 0 end ) as ncttl from share where "&_
      " ldate >= '"&mmdate&"'   "&_
      " group by memno ) c on a.memno=c.memno "&_
     "  where  "&styfield&" a.accode='"&memno&"'    and a.mstatus not in ('C','P','B','I' ) and a.wdate is null  "&_
     "   order by a.memno  "

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
<title>���~��</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">

<style type='text/css'>
p {page-break-after: always;}
</style>

</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">

<%
do while not rs.eof
	xmemno = rs("memno")
	memcname = rs("memcname")
	memname = rs("memname")
	B1=rs ("B1")
	B2=rs ("B2")
	B1relation = rs ("B1relation")
	B2relation = rs("B2relation")
	mstatus = rs("mstatus")
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

	SQl = "select lnnum,ldate  from loan  where memno='"&xmemno&"' and ldate>='"&xxdate&"' and ldate<='"&yydate&"' order by lnnum,ldate,right(code,1),left(code,1)  " 
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
		set ms = conn.execute("select *  from loan where  memno='"&xmemno&"'   and ldate<'"&yydate&"'   order by  memno,ldate,uid,right(code,1),left(code,1) " )
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
				if lnbal(1) <> "" and  yy = 0 then
                        bal = lnbal(1)
                        xx = xx + 1	
                end if
              	yy = 1
				xdate = right("0"&day(ms("ldate")),2)&"/"&right("0"&month(ms("ldate")),2)&"/"&year(ms("ldate"))
				
				select case ms("code")
           			case "0D"
						bal = ms("amount")
						lnbal(xx)=bal 
					case  "E0","E1","E2","E3","E4","E6","E7","EC"
						lnramt(xx) =ms("amount")
						bal = bal - ms("amount")
						lnbal(xx) = bal
					case  "ER"
						lnramt(xx) =ms("amount")
						bal = bal + ms("amount")
						lnbal(xx) = bal                
					case  "DE"
						lnramt(xx) =ms("amount") 
						lnbal(xx) = bal
					case  "D8"
						lnramt(xx) =ms("amount")
						bal = 0
						lnbal(xx) = 0        
				end select 
				
				if left(ms("code"),1) ="E" or ms("code")="0D" or ms("code")="D8"  or ms("code")="D9" OR ms("code")="DE" or ms("code")="NE" then
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
					xx = xx + 1          
					bb = bb + 1
				end if   
				
				if (ms("code")="F3"  or ms("code")="F1" or ms("code")="F2" or ms("code")="F7" or ms("code")="F0"   or rs("code")="FR" ) then 
					if ldate(xx-1) = ms("ldate") and xcode(xx-1) = "E"&right(ms("code"),1) then
							lniamt(xx-1) = ms("amount")
					end if
					
					if ldate(xx-1) <> ms("ldate") and (left(ms("code"),1) = "F" and code<>"F3" )  then
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
																																																						lniamt(xx) = ms("amount")
						lnbal(xx) = lnbal(xx-1) 
						lcode(xx) = LoanCode(ms("code")) 
						xx = xx + 1          
						bb = bb + 1
					end if
					
					if ms("code") = "F3" then
						xidate = ms("ldate")
						set mr1 = conn.execute("select count(*) from loan where memno='"&ms("memno")&"' and ldate = '"&xidate&"'  and code='F3' group by memno,ldate,code ")
						xcnt = mr1(0)
						do while xcnt > 0 and not mr1.eof
							lniamt(xx - xcnt)= ms("amount")
							xcnt = xcnt -1
							if xcnt > 0 then
								ms.movenext
							end if
						loop
					end if         
				end if

				if ms("code")="D9"  then            
					set ms1=conn.execute("select chequeamt,lnflag,appamt,loantype from loanrec where lnnum='"&ms("lnnum")&"'  ")
                    if  not ms1.eof then
                        loantype = ms1("loantype")   
                        lnflag = ms1(1)
                        if lnflag = "Y" then                        
                            if ms1(0) <> 0 then  
                                   lniamt(xx-1) =  ms1(0)
                            else
                                   lniamt(xx-1) =""
                            end if
                            if loantype ="N" then
								lcode(xx-1)  = "+ �s�U  ="
                            else
								lcode(xx-1)=" ������ "
                            end if               
                            lnbal(xx-1) = ms1(2)  
                            bal = lnbal(xx-1)          
                        else
                                 lnbal(xx-1) = ms1("appamt")
                                 lniamt(xx-1) =""
                                 bal = lnbal(xx-1)
                                 lcode(xx-1) ="�s�U"  
                        end if                    
                    end if
                    ms1.close 
                    MX = 0     
                end if 

				if  ms("code")="DF" or ms("code")="NF"   then
					mx = 0
					lniamt(xx-1) = ms("amount")    
					if ms("memno") ="4480"  and ms("code")="DF"  and  ymd(ms("ldate")) >="2016/07/30" and ms("lnnum")= "2013080003" then
						lnbal(xx-1) = bal + ms("amount") 
						bal =    lnbal(xx-1)
					end if 
				end if 
 			else

			if  left(ms("code"),1)="E" or ms("code")="D8"  then        
				lnbal(1)=lnbal(1)-ms("amount")
			else 
				if  ms("code")="DF" or ms("code")="NF"   then
					mx = 0
					lniamt(xx-1) = ms("amount")    
					if ms("memno") ="4480"  and ms("code")="DF"  and  ymd(ms("ldate")) >="2016/07/30" and ms("lnnum")= "2013080003" then
                      lnbal(xx-1) = bal + ms("amount") 
                    bal =    lnbal(xx-1)
				end if 
			end if 
			
            if trim(ms("code"))="0D" or trim(ms("code"))="D9"   then 
                if ms("code") = "D9" then
                  set rs1=conn.execute("select chequeamt,lnflag,appamt,loantype from loanrec where lnnum='"&ms("lnnum")&"'  ")
                      if  not rs1.eof then
                          lnbal(1) = rs1(2)                      
                      end if
                else                                                 
                     lnbal(1) = ms("amount")
                end if
           end if
        end if  
      
        lnnum(1) = ms("lnnum")
         lndate(1) = chkdate
        lcode(1) = "�U�ڵ��l" 
        lniamt(1)  =""     
        aa = 1
	end if
 
	if  not ms.eof then
		ms.movenext 
	end if
	loop
	ms.close
end if

if lniamt(xx) = "" and lnramt(xx) ="" then
   lnnum(xx) = ""
end if

bb = xx
yy = 0
xx = 1
aa = 0

set ms = conn.execute("select * from share where memno='"&xmemno&"' and ldate<'"&yydate&"'  order by memno,ldate,code ")
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
				if left(ms("code"),1)="B" or  ms("code")="MF" then
					mkdate2 = dateserial(year(ms("ldate")),month(ms("ldate")),1)
					if mkdate0 <= ms("ldate") then
						for i = 1 to 12
							if dvdate(i)= mkdate1  then
								divd(i) = sbal(xx)
						     end if
						next                      
				    end if
					'This option will change period balance to last balance change  
					for i = 1 to 12
						if dvdate(i)<= mkdate2 then
							divd(i) = sbal(xx)
						end if
					next 
				end if   
  
        case   "AI","CH"

                sbal(xx) = sbal(xx-1)  
                mkdate1 = dateserial(year(ms("ldate")),month(ms("ldate"))+1,1)
                for i = 1 to 12
                      if dvdate(i)= mkdate1 then
                       
                          divd(i) = sbal(xx)
                       end if 
                next 
   
        case  "A1","A2","A3","C0","C1","C3" ,"A0","A7" ,"A4" 
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
		if  left(ms("code"),1)="G" or left(ms("code"),1)="H" or left(ms("code"),1)="B" or ms("code")="MF" then        
			sbal(xx)=sbal(xx)-ms("amount")
		else 
			if ms("code")="0A" or left(ms("code"),1)="C" or left(ms("code"),1)="A" then 
				if ms("code")<>"AI"  then  
					sbal(xx) = sbal(xx) +ms("amount")
				end if
			end if  
			sdate(xx) = chkdate
			scode(xx) = "�Ѫ����l" 
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
<p>
<center>
<BR>
<font face="arial, helvetica, sans-serif" size="3" color="#336699"><b>���ȸp���u�x�W���U��</b></font><br>
<font face="arial, helvetica, sans-serif" size="1" color="#336699"><b>Water Supplies Department Staff Credit Union<br>Membership, Accounting, Savings and Loans Software</b></font>

<BR>
<font face="arial, helvetica, sans-serif" size="2" color="#336699"><b><%=Nyear%>�@���~��</b></font>
</center>
<table width="1060" border="0" cellspacing="0" cellpadding="0">
     <tr>
          <td width="400" align="left"><font size="3"  face="�з���" >�@�@<%=xmemno%>�@<%=memcname%></font></td>
          <td width="660" align="left">�@�@</td>
     </tr>
     <tr>
          <td width="400" align="left">�@�@ <%=memname%></font></td> 
          <td width="660" align="left"><font size="3"  face="�з���" >�@�@�@�@�]�p�����^<%=xmemcname%>�]�q�ܡ^<%=xmemContactTel%></font></td> 
     </tr>
     
</table>


<br>
<table border="0" cellspacing="1" cellpadding="1" align="center" >
	<tr >
		<td width=100 align="center"><font size="2"  face="�з���" >���</font></td> 
		<td width=80  align="center"><font size="2"  face="�з���" >�Ѫ�</font></td> 
		<td width=100 align="center"><font size="2"  face="�з���" >���l</font></td> 
		<td width=100 align="center"><font size="2"  face="�з���" >���O</font></td> 
        <td width=1   align="center"> </td>	
		<td width=100 align="center"><font size="2"  face="�з���" >���</font></td> 
		<td width=100 align="center"><font size="2"  face="�з���" >�Q��</font></td> 
		<td width=100 align="center"><font size="2"  face="�з���" >�C���ٴ�</font></td> 
		<td width=100 align="center"><font size="2"  face="�з���" >�s�U�`�B/<br>���l</font></td> 
		<td width=100 align="center"><font size="2"  face="�з���" >���O</font></td>            
	</tr>
	<tr><td colspan=11><hr></td></tr>
	<% xx = 1
    do while xx < smax
		<!-- Share Listing  -->
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

			
			<td width=100 align="center"><font size="2"  face="�з���" ><%=scode(xx)%></font></td>    
			<td  width=1 align="center"> </td>	
			
				
			
			<%if lnnum(xx)<>"" then%>
				<td width=100 align="center"><%=lndate(xx)%></td>
				
				<%If lcode(xx)="�U�ڵ��l" then%> 
					<td ></td>
					<td ></td>
					<td  align="right"> <%=formatnumber(lnbal(xx),2)%></td>
				<%end if%> 
				<%If lcode(xx)="�w����b" then%> 
					<td  align="right"><%=formatnumber(lniamt(xx),2)%></td>  <!--�Q�� -->
					<td  align="right"><%=formatNumber(lnramt(xx),2)%></td>
					<td  align="right"> <%=formatnumber(lnbal(xx),2)%></td>
				<%end if%> 
				<%if lcode(xx)="+ �s�U  =" then %>                      
					<td  align="right"><%=formatnumber(lniamt(xx),2)%></td>  <!--�Q�� -->
					<td width=100 align="center"><font size="2"  ><%=lcode(xx)%></font></td> 			
					<td  align="right"><%=formatNumber(lnbal(xx),2)%></td>
				<%end if%>
				<%if lcode(xx)="�w�в��" then%>  
						<td  align="right"><%=formatnumber(lniamt(xx+1),2)%></td>  <!--�Q�� -->
						<td  align="right"><%=formatNumber(lnramt(xx+1),2)%></td>
						<td  align="right"> <%=formatNumber(lnbal(xx+1),2)%></td>
				<%end if%> 
				<%if lcode(xx)="�Ȧ���" then%>  
						<td ></td>  <!--�Q�� -->
						<td  align="right"><%=formatNumber(lnramt(xx+1),2)%></td>
						<td  align="right"> <%=formatNumber(lnbal(xx+1),2)%></td>
				<%end if%> 
				
				<%if lcode(xx)="�{���ٴ�"   then%>
						<%if  lcode(xx-1)="�U�ڵ��l" then%>
							<td  align="right"><%=formatnumber(lniamt(xx),2)%></td>  <!--�Q�� -->
							<td  align="right"><%=formatNumber(lnramt(xx),2)%></td>
							<td  align="right"> <%=formatNumber(lnbal(xx),2)%></td>
						<%end if%>
					
						<%if  lcode(xx-1)="�w�в��" then%>
							<td  align="right"><%=formatnumber(lniamt(xx),2)%></td>  <!--�Q�� -->
							<td  align="right"><%=formatNumber(lnramt(xx),2)%></td>
							<td  align="right"> <%=formatNumber(lnbal(xx),2)%></td>
						<%end if%>
						
						<%If lcode(xx-1)="�Ȧ���" then%> 
							<td ></td>  <!--�Q�� -->
							<td  align="right"><%=formatNumber(lnramt(xx),2)%></td>
							<td  align="right"> <%=formatnumber(lnbal(xx),2)%></td>
						<%end if%>
						<%If lcode(xx-1)="�Ȧ���b" then%> 
							<td  align="right"><%=formatnumber(lniamt(xx+1),2)%></td>  <!--�Q�� -->
							<td  align="right"><%=formatNumber(lnramt(xx+1),2)%></td>
							<td  align="right"> <%=formatnumber(lnbal(xx),2)%></td>
						<%end if%>
				
				<%end if%>  
 				<%If lcode(xx)="�U�ڲM��" then%> 
					<td  align="right"><%=formatnumber(lniamt(xx),2)%></td>  <!--�Q�� -->
					<td  align="right"><%=formatNumber(lnramt(xx),2)%></td>
					<td  align="right"> <%=formatnumber(lnbal(xx),2)%></td>
				<%end if%>

				<%If lcode(xx)="�Ȧ���b" then%> 
					<%if xmemno=3331 then%>
						<td></td> 
					<%else%>
						<td  align="right"><%=formatNumber(lniamt(xx),2)%></td>
					<%end if%>	
					<td  align="right"><%=formatNumber(lnramt(xx),2)%></td>
					<td  align="right"> <%=formatnumber(lnbal(xx),2)%></td>
				<%end if%> 
				<%if lcode(xx)="+ �s�U  =" then %>  
					<td></td> 
				<%else%>
					<td width=100 align="center"><font size="2" face="�з���"  ><%=lcode(xx)%></font></td> 
				<%end if%>
		</tr>   
		<% 
		end if


        xx = xx + 1
        line = line + 1
    loop
    if smax<21 then
        uu = 21 - smax 
        for i = 1 to uu %>
			<tr><td colspan=11>�@�@�@�@�@</td></tr>
			<%
			line = line + 1
		next
    end if   %>
</table>

<table border="0" cellspacing="1" cellpadding="1" >
<tr>
<td width="340" valign="top">
<%if xlnnum<>"" then%>      
	<table border="0" cellspacing="1" cellpadding="1" align="center" >
	<tr>
         <td width="100" >�ɾڽs��</td>
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
	guname1 = "*****"
	guname2 = "*****"
	guname3 = "*****"
	guid1   = "*****"
	guid2   = "*****"
	guid3   = "*****"
	guln1   = "*****"
	guln2   = "*****"
	guln3   = "*****"
	guoid1  = "*****"
	guoid2  = "*****"
	guoid3  = "*****"
	Set es = server.createobject("ADODB.Recordset")
	sqlstr = " select * from guarantor where lnnum = '"&xlnnum&"' "
	es.open sqlstr, conn, 1, 1
	xx = 1
	do while  not es.eof 
			select case xx
				case 1 
					guid1 = es("guarantorID")                
					guname1 = es("guarantorName")
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
									 if rs1("code")<>"AI" then
										bal = bal+ rs1("amount")
									 end if
							end select
							rs1.movenext
						loop
						rs1.close
						gusave1 = bal
					end if
				case 2
					guid2 = es("guarantorID")               
					guname2 = es("guarantorName")
					if guid2 <> "" then
						bal = 0
						sql1 = " select * from share where memno = "& guid2                  
						rs1.open sql1, conn, 2, 2
						do while not rs1.eof
							select case left(rs1("code"),1)
								case "G","H","B","M"
									 bal = bal - rs1("amount")
								case "A","T","C","0"
									 if rs1("code")<>"AI" then
										bal = bal+ rs1("amount")
									 end if
							end select
							rs1.movenext
						loop
						rs1.close
						gusave2 = bal
					end if
				case 3
					guid3 = es("guarantorID")            
					guname3 = es("guarantorName")
					if guid3 <> "" then
						bal = 0
						sql1 = " select * from share where memno = "& guid3                  
						rs1.open sql1, conn, 2, 2
						do while not rs1.eof
							select case left(rs1("code"),1)
								case "G","H","B","M"
									 bal = bal - rs1("amount")
								case "A","T","C","0"
									if rs1("code")<>"AI" then
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
			es.movenext   
	loop
	es.close
	xx = 1
	Set ms = server.createobject("ADODB.Recordset")
	sql1 = "select * from guarantor  where  guarantorID ='"&xlnnum&"' "
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
		<tr>
			<td width="50">&nbsp;</td>
			<td width="100">1.��O�H<td>
			<td width="50"><%=guid1%></td>
			<td width="200"><%= guname1%></td>
		</tr>
		<tr>
			<td width="50">&nbsp;</td>
			<td width="100">2.��O�H <td>
			<td width="50"><%=guid2%></td>
			<td width="200"><%=guname2%></td>
		</tr>
		<tr>
			<td width="50">&nbsp;</td>
			<td width="100">3.��O�H <td>
			<td width="50"><%=guid3%></td>
			<td width="200"><%=guname3%></td>
		</tr>
		<tr>
			<td width="50">&nbsp;</td>
			<td width="100">1.��O��L�H <td>
			<td width="50"><%=guoid1%></td>
			<td width="300"><%=  guln1%></td>
		</tr>
		<tr>
			<td width="50">&nbsp;</td>
			<td width="100">2.��O��L�H <td>
			<td width="50"><%=guoid2%></td>
			<td width="200"><%=  guln2%></td>
		</tr>
		<tr>
			<td width="50">&nbsp;</td>
			<td width="100">3.��O��L�H <td>
			<td width="50"><%=guoid3%></td>
			<td width="200"><%=  guln3%></td>
		</tr>
	</table>
	</td>
		 
	<% 
	line = line + 6 
end if%>

	<td width="550" valign="top">
	<cemter>
	<font size="2"  face="�з���" ><b>�Ѯ��p���</b></font><br>
	</center>
        <table border="1" cellpadding="0" cellspacing="0"  >
        <tr>
            <td width="120" align="center">�@�@����@�@</td>
            <td width="200" align="center"  >�@�@���e�@�@</td>
            <td width="100" align="center" >�@�@���B�@�@</td>
        </tr>
		<%
        ttlamt = 0
	for i = 1 to 12
				sldate = "01/"&right("0"&month(dvdate(i)),2)&"/"&year(dvdate(i))
				if divd(i)= 0 then
							divd(i)= divd(i-1)
				end if
				if right(round(divd(i)),2)="26" and divd(i-1)=0 then divd(i)=divd(i)-26   end if
				if right(round(divd(i)),2)="36" and divd(i-1)=0 then divd(i)=divd(i)-36   end if
				if  divd(i) <>"" then
					'***********************************************************			
					if rs("memno")=2871 and sldate="01/11/2018"  then
								divd(i)= 11595.45
					end if
					if rs("memno")=2871 and sldate="01/10/2018"  then
								divd(i)= 10095.45
					end if
					if rs("memno")=2871 and sldate="01/09/2018"  then
								divd(i)= 8572.20
					end if
					if rs("memno")=2871 and sldate="01/08/2018"  then
								divd(i)= 7072.20
					end if	
					if rs("memno")=2871 and sldate="01/07/2018"  then
								divd(i)= 5572.20
					end if	
					
					if rs("memno")=3300 and sldate="01/09/2018"  then
								divd(i)= 2873.52
					end if	
					if rs("memno")=3300 and sldate="01/08/2018"  then
								divd(i)= 2573.52
					end if	
					if rs("memno")=3300 and sldate="01/07/2018"  then
								divd(i)= 2573.52
					end if	
					if rs("memno")=3341 and sldate="01/11/2018"  then
								divd(i)= 420.23
					end if	
					if rs("memno")=3341 and sldate="01/10/2018"  then
								divd(i)= 420.23
					end if	
					if rs("memno")=3341 and sldate="01/09/2018"  then
								divd(i)= 420.23
					end if	
					if rs("memno")=3341 and sldate="01/08/2018"  then
								divd(i)= 420.23
					end if	
					if rs("memno")=3341 and sldate="01/07/2018"  then
								divd(i)= 420.23
					end if
					if rs("memno")=3391 and sldate="01/09/2018"  then
								divd(i)= 1441.13
					end if	
						
					if rs("memno")=3391 and sldate="01/08/2018"  then
								divd(i)= 1441.13
					end if	
					if rs("memno")=3391 and sldate="01/07/2018"  then
								divd(i)= 1441.13
					end if	
					if rs("memno")=3770 and sldate="01/10/2018"  then
								divd(i)= 2344.53
					end if	
					if rs("memno")=3770 and sldate="01/09/2018"  then
								divd(i)= 1844.53
					end if	
					if rs("memno")=3770 and sldate="01/08/2018"  then
								divd(i)= 344.53
					end if
					if rs("memno")=3770 and sldate="01/07/2018"  then
								divd(i)= 344.53
					end if
					if rs("memno")=4119 and sldate="01/07/2018"  then
								divd(i)= 755.24
					end if
					if rs("memno")=4119 and sldate="01/08/2018"  then
								divd(i)= 755.24
					end if
					if rs("memno")=4372 and sldate="01/09/2018"  then
								divd(i)= 1783.94
					end if	
					if rs("memno")=4372 and sldate="01/08/2018"  then
								divd(i)= 1783.94
					end if
					if rs("memno")=4372 and sldate="01/07/2018"  then
								divd(i)=1783.94
					end if
					if rs("memno")=4388 and sldate="01/12/2018"  then
								divd(i)= 9728.39 
					end if					
					if rs("memno")=4388 and sldate="01/11/2018"  then
								divd(i)= 8228.39
					end if	
					if rs("memno")=4388 and sldate="01/10/2018"  then
								divd(i)=  6728.39
					end if	
					if rs("memno")=4388 and sldate="01/09/2018"  then
								divd(i)= 6728.39
					end if	
					if rs("memno")=4388 and sldate="01/08/2018"  then
								divd(i)= 6728.39
					end if	
					if rs("memno")=4388 and sldate="01/07/2018"  then
								divd(i)= 6728.39
					end if
					if rs("memno")=4442 and sldate="01/07/2018"  then
								divd(i)= 4180.24
					end if
					if rs("memno")=4508 and sldate="01/12/2018"  then
								divd(i)= 10371.80
					end if

					if rs("memno")=4508 and sldate="01/11/2018"  then
								divd(i)= 10371.80
					end if
					if rs("memno")=4508 and sldate="01/10/2018"  then
								divd(i)= 9871.80
					end if
					if rs("memno")=4508 and sldate="01/09/2018"  then
								divd(i)= 9871.80
					end if
					if rs("memno")=4508 and sldate="01/08/2018"  then
								divd(i)= 9871.80
					end if	
					if rs("memno")=4508 and sldate="01/07/2018"  then
								divd(i)= 9871.80
					end if	
					if rs("memno")=4515 and sldate="01/01/2019"  then
								divd(i)= 3816.55
					end if

				    if rs("memno")=4515 and sldate="01/12/2018"  then
								divd(i)=  3816.55
					end if

					if rs("memno")=4515 and sldate="01/11/2018"  then
								divd(i)=  3816.55
					end if
					if rs("memno")=4515 and sldate="01/10/2018"  then
								divd(i)=  3816.55
					end if
					if rs("memno")=4515 and sldate="01/09/2018"  then
								divd(i)=  3816.55
					end if
					if rs("memno")=4515 and sldate="01/08/2018"  then
								divd(i)=  3816.55
					end if	
					if rs("memno")=4515 and sldate="01/07/2018"  then
								divd(i)= 3800.19
					end if	

					if rs("memno")=4556 and sldate="01/07/2018"  then
								divd(i)= 369.95
					end if	
					if rs("memno")=4629 and sldate="01/08/2018"  then
								divd(i)= 559.00
					end if	
					if rs("memno")=4629 and sldate="01/07/2018"  then
								divd(i)= 3800.19
								
								'*** 2020  ****
					end if	
					if rs("memno")=4851 and sldate="01/07/2020"  then
								divd(i)= 9027.50
					end if
					if rs("memno")=4851 and sldate="01/08/2020"  then
								divd(i)= 9627.50
					end if
					if rs("memno")=4851 and sldate="01/09/2020"  then
								divd(i)= 10227.50
					end if				
					if rs("memno")=4851 and sldate="01/10/2020"  then
								divd(i)= 10828.27
					end if					
					if rs("memno")=4851 and sldate="01/11/2020"  then
								divd(i)= 11483.27
					end if					
					if rs("memno")=4851 and sldate="01/12/2020"  then
								divd(i)= 12083.27
					end if					
					if rs("memno")=4851 and sldate="01/01/2021"  then
								divd(i)= 12683.27
					end if					
					if rs("memno")=4851 and sldate="01/02/2021"  then
								divd(i)= 13283.30
					end if					
					if rs("memno")=4851 and sldate="01/03/2021"  then
								divd(i)= 13883.30
					end if					
					'****** 4986  *******
					if rs("memno")=4986 and sldate="01/07/2020"  then
								divd(i)= 2500
					end if
					if rs("memno")=4986 and sldate="01/08/2020"  then
								divd(i)= 2500
					end if
					if rs("memno")=4986 and sldate="01/09/2020"  then
								divd(i)= 2500
					end if
					if rs("memno")=4986 and sldate="01/10/2020"  then
								divd(i)= 2500
					end if					
					if rs("memno")=4986 and sldate="01/11/2020"  then
								divd(i)= 2500
					end if					
					if rs("memno")=4986 and sldate="01/12/2020"  then
								divd(i)= 2500
					end if					
					if rs("memno")=4986 and sldate="01/01/2021"  then
								divd(i)= 2500
					end if					
					if rs("memno")=4986 and sldate="01/02/2021"  then
								divd(i)= 2500
					end if					
					if rs("memno")=4986 and sldate="01/03/2021"  then
								divd(i)= 2500
					end if				
					if rs("memno")=4986 and sldate="01/04/2021"  then
								divd(i)= 2500
					end if
					if rs("memno")=4986 and sldate="01/05/2021"  then
								divd(i)= 2500
					end if

					
	
					'***********************************************************************	
					desc   = "�@"&formatNumber(divd(i))&" X "&rate&"/"&"100"&"/12"
				else
					desc   = "�@0 X"&rate&"/"&"100"&"/12"
				end if
           xamt   =formatnumber( divd(i)*rate/100/12,2)*1
           ttlamt = ttlamt + xamt%>
       	
		 <tr>   
            <td width="100" align="center"><%=sldate%>�e</td>
            <td width="200" align="left"  ><%=desc%></td>
            <td width="100" align="right" ><%=formatnumber(xamt,2)%></td>
        </tr>
		<%next%>
        <tr>
            
            <td width="100" align="center">&nbsp;</td>
            <td width="200" align="right"  >�X�@���B�G</td>
            <td width="100" align="right" ><%=formatnumber(ttlamt,2)%></td>
        </tr>

        </table>
	</td>
	</tr>
	</table>

	<BR>
	<font face="arial, helvetica, sans-serif" size="3" ><b>�@�@�@�@</b></font><br>
	<font face="arial, helvetica, sans-serif" size="5" ><b>�@�@�@�@���~�H 1. <%=B1%> ,   ���~�H 2. <%=B2%> </b></font><br>
	<font face="arial, helvetica, sans-serif" size="3" ><b>�@�@�@�@</b></font><br>
	<font face="arial, helvetica, sans-serif" size="3" ><b>�@�@�@�@</b></font><br>
	<font face="arial, helvetica, sans-serif" size="3" ><b>�@�@�@�@�� - (1) ���~�ת���|�O�N�ѥ�����I </b></font><br>
	<%if ttlamt > 100 then %>
		<%if mstatus ="M" or mstatus="T" then %>
			<font face="arial, helvetica, sans-serif" size="3" ><b>�@�@�@�@�@ - (2) �դU�󥻦~�ש��Ȩ����Ѯ���<%=formatnumber(ttlamt,2)%>���A�H�䲼��I�C</b></font><br>     
		<%else%>
			<font face="arial, helvetica, sans-serif" size="3" ><b>�@�@�@�@�@ - (2) �դU�󥻦~�ש��Ȩ����Ѯ���<%=formatnumber(ttlamt,2)%>���A��<%=pyear%>�~<%=int(mon)%>��<%=nday%>��۰����J�դU�Ȧ��f�C</b></font><br>
		<%end if%>
	<%else%>
		<font face="arial, helvetica, sans-serif" size="3" ><b>�@�@�@�@�@ - (2) �դU�󥻦~�תѮ���<%=formatnumber(ttlamt,2)%>���A��<%=pyear%>�~<%=int(mon)%>��<%=nday%>��s�J�դU���Ѫ����C</b></font><br>
	<%end if%>

	<font face="arial, helvetica, sans-serif" size="3" ><b>�@�@�@�@�@ - (3) �p�դU�Q���������̷s�̧֪������A�Хߨ�w��WhatsApp�{���A�ñN�q�ܸ��X 5476 6906 </b></font><br>
	<font face="arial, helvetica, sans-serif" size="3" ><b>�@�@�@�@�@  �@�@�[�J�դU���q��ï���Y�i�C�P�ɥ�i�H�ɵn�J�������} www.wsdscu.org �s���Ҧ��q�i�γ̷s�����C</b></font><br>
	<font face="arial, helvetica, sans-serif" size="3" ><b>�@�@�@�@�@ - (4) ���T�����O�A�����N���A�l�H���ʳq�i���h������A�ӧO�h������p���ݭn�A�i�p������¾���C</b></font><br>
	<BR>
	<BR>
	<BR>
<%for i = 1 to 12
           divd(i) = 0
  next 
       
       rs.movenext
       loop
%>

</table>



</center>
</p>
</body>
</html>
<%

rs.close
set rs=nothing
conn.close
set conn=nothing
%>
