<!-- #include file="../conn.asp" -->
<!-- #include file="init.asp" -->
<%
Function ShareCode(ByVal x)
	select case x
			case "0A"
				ShareCode = "股金結餘"
			case "A0"
				ShareCode = "退還貸款"
			case "A1"
				ShareCode = "銀行轉帳"
			case "A2"
				ShareCode ="庫房轉帳"
			case "A3"
				ShareCode ="現金存款"
			case "A4"
				ShareCode ="保險金"
			case "B0"
				ShareCode="股金還款"
			case "A7"
			
			case "B1"
					ShareCode="退股"
			CASE "AI"
					ShareCode ="脫期　　" 
			CASE "D9"
					ShareCode ="新貸銀行"  
			CASE "B0"
					ShareCode ="現金退股"
			case "B3"
					ShareCode ="退還現金"
					
			case "C0"
					ShareCode="股息　　"
			case "CH"
					ShareCode="暫停股息"        
			case "C1"
					ShareCode="股息銀行" 
			case "C3"
					ShareCode="股息現金" 
 			case "G0","G1","G2","G3"
					ShareCode = "入社費"
			case "E7"	
					ShareCode ="調(+)整"
			case "H7"
					ShareCode ="調(-)整"	
			case "H0","H1","H2","H3"
					ShareCode = "協會費" 
			case "MF"
					ShareCode = "冷戶費" 
	end select
End Function

Function LoanCode(ByVal x)
	select case x
        case "0D"
				LoanCode	="貸款結餘"
        case "E1"
				LoanCode	= "銀行轉帳"
        case "E2"
				LoanCode	="庫房轉帳"
        case "EC"
                LoanCode	="劃消金額"
        case "E3"
				LoanCode	="現金還款"
        case "E0"
            if ms("amount") > 0 then
				LoanCode	="股金還款"
            else
				LoanCode	="退還本息"
            end if 
        case "F0"
            if ms("amount") > 0 then
				LoanCode	="股金還款"
            else
				LoanCode	="退還利息"
               end if 
        case "E6"
                LoanCode	="退款"
        case "E7"
                LoanCode	="調整"
        case "F1"
                LoanCode	="銀行還息"
        case "F2"
                LoanCode	="庫房還息"
        case "F3"
                LoanCode	="現金還息"
        case "ER"
				LoanCode	="退還本金"
        case "F3"
                LoanCode	="現金還息"  
        case "FR"
				LoanCode	="退還利息"
        CASE "DE"
				LoanCode	="銀行脫期" 
        CASE "DF"
				LoanCode	="利息脫期" 
        CASE "NE"
				LoanCode	="庫房脫期"            
        CASE "D8"
                LoanCode	="貸款清數"
        case "DE","NE"
				mx 			= 0  
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
mon 	= mid( request.form("paidday"),4,2)
nday	= left( request.form("paidday"),2)
xname	= request.form("accode")
rate  	= request.form("nrate")
pos 	= instr(xname,"-")
if pos>0 then
	memno = mid(xname,1,4)
	mname =  mid(xname,pos+1,50)
else
	response.redirect "fyPprt.asp"
end if
if memno<>"9999"  then
	memno 	= left(xname,pos-1)
	mname 	=  mid(xname,pos+1,50)
	styfield = "accode='"&memno&"' and "
else
	styfield = "accode='9999' and "
if instr(mname,"(1-500)") > 0 then
	styfield = styfield & "(a.memno>='1' and a.memno<='500') "
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
if instr(mname,"(3001-4000)") > 0 then
      styfield =styfield &  " (a.memno >='3001' and   a.memno<='4000') and "
else
if instr(mname,"(4001-5000)") > 0 then
      styfield =styfield &  " (a.memno >='4001' and   a.memno<='5000') and "
else
if instr(mname,"(5001-6000)") > 0 then
      styfield =styfield &  " (a.memno >='5001' and   a.memno<='6000') and "
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
   xmemcname = "工作人員"
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

SQl = "select memno,memname,memcname,mstatus  from memmaster where  "&styfield&"  mstatus not in ('C','P','B' ) order by memno   "
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
'''response.write(sql)
'''response.end
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

<style type='text/css'>
	p {page-break-after: always;}
</style>
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
<%
do while not rs.eof
    xmemno 		= rs("memno")
    memcname 	= rs("memcname")
    memname 	= rs("memname")
    mstatus 	= rs("mstatus")
	B1			= rs ("B1")
	B2			= rs ("B2")
	B1relation  = rs ("B1relation")
	B2relation  = rs("B2relation")	
	
	
    line 		= 8
    for i = 1 to 50
         sdate(i) 	= ""
         scode(i) 	= "" 
         samt(i) 	= ""
         sbal(i)  	= ""
         lnnum(i) 	= ""
         lndate(i) 	= ""
         lncode(i)  = ""
         xcode(i)  	= ""
         lnramt(i) 	= ""
         lniamt(i) 	= ""
         lnbal(i)  	= ""
         ldate(i)  	= ""
         lcode(i) 	= "" 
    next  
    xlnnum	= ""
    ylnnum 	= ""
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
        cc 	= 	0
        xx 	= 	1
        qx	= 	0 
        MX	=	0
        zero= 	0
        bb 	= 	0
        yy 	= 	0
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
					case  "E0","E1" , "E4" , "E2" , "E3" , "E6" ,"E7","EC"
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
						lnnum(xx) 	= ms("lnnum")
						newln 		= 0
						xyear 		= year(ms("ldate"))
						xmon  		= month(ms("ldate"))
						xday  		= day(ms("ldate"))
						oldamt 		= 0
						xcode(xx) 	= ms("code")
						ldate(xx) 	= ms("ldate")
						lndate(xx) 	= right("0"&xday,2)&"/"&right("0"&xmon,2)&"/"&xyear
						xdate      	= right("0"&xday,2)&"/"&right("0"&xmon,2)&"/"&xyear     
						lcode(xx) 	= LoanCode(ms("code")) 
						xx			= xx + 1          
						bb 			= bb + 1
				end if 
				
				if (ms("code")="F3"  or ms("code")="F1" or ms("code")="F2" or ms("code")="F7" or ms("code")="F0"   or rs("code")="FR" ) then 
					if ldate(xx-1) = ms("ldate") and xcode(xx-1) = "E"&right(ms("code"),1) then
						lniamt(xx-1) = ms("amount")
					end if
					if ldate(xx-1) <> ms("ldate") and (left(ms("code"),1) = "F" and code<>"F3" )  then
						lnnum(xx) 	= ms("lnnum")
						newln 		= 0
						xyear 		= year(ms("ldate"))
						xmon  		= month(ms("ldate"))
						xday  		= day(ms("ldate"))
						oldamt 		= 0
						xcode(xx) 	= ms("code")
						ldate(xx) 	= ms("ldate")
						lndate(xx)	= right("0"&xday,2)&"/"&right("0"&xmon,2)&"/"&xyear
						xdate     	= right("0"&xday,2)&"/"&right("0"&xmon,2)&"/"&xyear    
						lniamt(xx)	= ms("amount")
						lnbal(xx) 	= lnbal(xx-1) 
						lcode(xx) 	= LoanCode(ms("code")) 
						xx 			= xx + 1          
						bb 			= bb + 1
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
                                lniamt(xx-1)= ""
                            end if
                            if loantype 	= "N" then
                               lcode(xx-1)  = "+ 新貸  ="
                            else
                               lcode(xx-1)	= " 更改期數 "
                            end if               
                            lnbal(xx-1) 	= ms1(2)  
							bal 			= lnbal(xx-1)          
						else
                            lnbal(xx-1) 	= ms1("appamt")
                            lniamt(xx-1)	= ""
                            bal 			= lnbal(xx-1)
                            lcode(xx-1) 	= "新貸"  
                        end if                     
                    end if
                    ms1.close 
                    mx = 0     
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
      
				lnnum(1) 	= ms("lnnum")
				lndate(1) 	= chkdate
				lcode(1) 	= "貸款結餘" 
				lniamt(1)  	= ""     
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
'****share DIS*****
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
        case	"0A"
                sbal(xx)=sbal(xx-1)+ms("amount")
                shareBalanceDate= dateserial(year(ms("ldate")),month(ms("ldate"))+1,1)
				for i = 1 to 12
						if dvdate(i)= shareBalanceDate then
							divd(i) = sbal(xx)
						end if 
				next 
				
		case	"E7"
                sbal(xx)=sbal(xx-1)+ms("amount")
		case	"H7"
                sbal(xx)=sbal(xx-1)-ms("amount")
				
        case 	"B1","MF"
				sbal(xx)=sbal(xx-1)-ms("amount")
				shareCurrDate  = dateserial(year(ms("ldate")),month(ms("ldate")),1)
				if  shareBalanceDate <= ms("ldate") then
						for i = 1 to 12
							if dvdate(i)<= shareBalanceDate then
								divd(i) = sbal(xx)
							end if
						next                      
					end if
					'This option will change period balance to last balance change  
					for i = 1 to 12
						if dvdate(i)<= shareCurrDate  then
							divd(i) = sbal(xx)
						end if
				next      
				
				
				

        case 	"G0" ,"H0","B0","B3","BE","BF","G3","H3" 
				sbal(xx)=sbal(xx-1)-ms("amount")
				if 	left(ms("code"),1)="B" or  ms("code")="MF" then
					shareCurrDate  = dateserial(year(ms("ldate")),month(ms("ldate")),1)
					if shareBalanceDate  <= ms("ldate") then
						for i = 1 to 12
							if dvdate(i)= shareBalanceDate then
								divd(i) = sbal(xx)
							end if
						next                      
					end if
					'This option will change period balance to last balance change  
					for i = 1 to 12
						if dvdate(i)<= shareCurrDate  then
							divd(i) = sbal(xx)
						end if
					next      
				end if   
        case  	"AI","CH"
                sbal(xx) = sbal(xx-1)  
                shareBalanceDate = dateserial(year(ms("ldate")),month(ms("ldate"))+1,1)
                for i = 1 to 12
                      if dvdate(i)= shareBalanceDate then
                          divd(i) = sbal(xx)
                       end if 
                next 
		case 	"A1","A2","A3","C0","C1","C3" ,"A0","A7" ,"A4" 
                sbal(xx) = sbal(xx-1) + ms("amount")
				shareBalanceDate = dateserial(year(ms("ldate")),month(ms("ldate"))+1,1)
                shareDate0  = ms("ldate")
                for i = 1 to 12
                    if dvdate(i)= shareBalanceDate then
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

      
	xx = 1
	if  left(ms("code"),1)="G" or left(ms("code"),1)="H" or left(ms("code"),1)="B" or ms("code")="MF" then        
        sbal(xx)=sbal(xx)-ms("amount")
	end if	 

	if aa > bb then
	   smax = aa
	else
	   smax = bb
	end if
	if sbal(1) = ""  then
	   sbal(1) = ""
	   scode(1)= ""
	end if

%>
<p>
<center>
	<BR>
	<font face="arial, helvetica, sans-serif" size="3" color="#336699"><b>水務署員工儲蓄互助社</b></font><br>
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
		<td width=80 align="center"> <font size="2"  face="標楷體" >股金</font></td> 
		<td width=100 align="center"><font size="2"  face="標楷體" >結餘</font></td> 
		<td width=100 align="center"><font size="2"  face="標楷體" >類別</font></td> 
        <td width=1 align="center"> </td>	
		<td width=100 align="center"><font size="2"  face="標楷體" >日期</font></td> 
		<td width=100 align="center"><font size="2"  face="標楷體" >利息</font></td> 
		<td width=100 align="center"><font size="2"  face="標楷體" >每月還款</font></td> 
		<td width=100 align="center"><font size="2"  face="標楷體" >新貸總額/<br>結餘</font></td> 
 		<td width=100 align="center"><font size="2"  face="標楷體" >類別</font></td>            
	</tr>
	<tr><td colspan=11><hr></td></tr>
	<%
    xx = 1
    do while xx < smax	%>  
        <tr>
			<td width=100 align="center"><%=sdate(xx)%></td>
			<%
			if samt(xx)<>"" then%>               
				<td width=80 align="right"><%=formatnumber(samt(xx),2)%></td>
				<%
			else%>
				<td width=80 align="right"></td>
				<%
			end if%>
			<%
			if sbal(xx)<>"" then
				if sbal(xx)>0 then %> 
					<td width=100 align="right"><%=formatnumber(sbal(xx),2)%></td>
					<%
				else%>
					<td width=100 align="right">&nbsp;</td>
					<%
				end if%>
				<%
			end if%>                            
			<td width=100 align="center"><font size="2"  face="標楷體" ><%=scode(xx)%></font></td>    
			<td width=1 align="center"> </td>	
			<%
			if lnnum(xx)<>"" then%>
				<td width=100 align="center"><%=lndate(xx)%></td>
				<%
				if lniamt(xx)<>"" then%>                
					<td  align="right"><%=formatnumber(lniamt(xx),2)%></td>
					<%
				else
					%>
					<td></td>
					<%
				end if%>
				<%
				if lnramt(xx) <> ""   then %>           
					<td  align="right"><%=formatNumber(lnramt(xx),2)%></td>
					<%
				else
					if lcode(xx)="+ 新貸  =" or lcode(xx)="更改期數 =" then
						'   if loantype = "E" then
						'      lcode(xx)=" 更改期數 "
						'   end if
						%>                      
						<td ><%=lcode(xx)%></td>
						<% 
						lcode(xx)=""
					else
						%>
						<td></td>
						<%                
					end if
				end if%>                                 
				<%
				if lnbal(xx)<>""   then%>                 
					<td  align="right"><%=formatnumber(lnbal(xx),2)%></td>
					<%
				else%>
					<td>&nbsp;</td>
					<%
				end if%>
				<td width=100 align="center"><font size="2"  face="標楷體" ><%=lcode(xx)%></font></td>         
				</tr>   
				<%
			end if
        xx = xx + 1
        line = line + 1
    loop
    if smax<21 then
         uu = 21 - smax 
       
  
       
         for i = 1 to uu
            
%>
          <tr><td colspan=11>　　　　　</td></tr>
<%
     
      line = line + 1
      next
      end if   

%>
  </table>

  

<table border="0" cellspacing="1" cellpadding="1" >
<tr>
<td width="340" valign="top">
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
	guname1 = "*****"
	guname2 = "*****"
	guname3 = "*****"
	 guid1 = "*****"
	 guid2 = "*****"
	 guid3 = "*****"
	 guln1 = "*****"
	 guln2 = "*****"
	 guln3 = "*****"
	 guoid1 = "*****"
	 guoid2 = "*****"
	 guoid3 = "*****"
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
			<tr>
				 <td width="50">&nbsp;</td>
				 <td width="100">1.擔保人<td>
				 <td width="50"><%=guid1%></td>
				 <td width="200"><%= guname1%></td>
			</tr>
			<tr>
				 <td width="50">&nbsp;</td>
				 <td width="100">2.擔保人 <td>
				 <td width="50"><%=guid2%></td>
				 <td width="200"><%=guname2%></td>
			</tr>
			<tr>
				 <td width="50">&nbsp;</td>
				 <td width="100">3.擔保人 <td>
				 <td width="50"><%=guid3%></td>
				 <td width="200"><%=guname3%></td>
			</tr>
			<tr>
				 <td width="50">&nbsp;</td>
				 <td width="100">1.擔保其他人 <td>
				 <td width="50"><%=guoid1%></td>
				 <td width="300"><%=  guln1%></td>
			</tr>
			<tr>
				 <td width="50">&nbsp;</td>
				 <td width="100">2.擔保其他人 <td>
				 <td width="50"><%=guoid2%></td>
				 <td width="200"><%=  guln2%></td>
			</tr>
			<tr>
				 <td width="50">&nbsp;</td>
				 <td width="100">3.擔保其他人 <td>
				 <td width="50"><%=guoid3%></td>
				 <td width="200"><%=  guln3%></td>
			 </tr>
	</table>
	</td>
		 
	<% 
	line = line + 6 
end if
      


%>

<td width="550" valign="top">
<cemter>
<font size="2"  face="標楷體" ><b>　　　　股息計算表</b></font><br>
</center>
        <table border="1" cellpadding="0" cellspacing="0"  >
        <tr>
            <td width="120" align="center">　　日期　　</td>
            <td width="200" align="center">　　內容　　</td>
            <td width="100" align="center">　　金額　　</td>
        </tr>
<%
       
        ttlamt = 0
        for i = 1 to 12
			sldate = "01/"&right("0"&month(dvdate(i)),2)&"/"&year(dvdate(i))
			if divd(i)= 0 then
				divd(i)= divd(i-1)
			end if
				'if right(round(divd(i)),2)="26"  then divd(i)=divd(i)-26   end if
				if rs("memno")=>4934 and right(round(divd(i)),2)="36" and divd(i-1)=0 then divd(i)=divd(i)-36   end if
				if  divd(i) <>"" then	   
				'***********************************************************			
					if rs("memno")=2871 and sldate="01/11/2018"  then	divd(i)= 11595.45	end if
					if rs("memno")=2871 and sldate="01/10/2018"  then	divd(i)= 10095.45	end if
					if rs("memno")=2871 and sldate="01/09/2018"  then	divd(i)= 8572.20	end if
					if rs("memno")=2871 and sldate="01/08/2018"  then	divd(i)= 7072.20	end if	
					if rs("memno")=2871 and sldate="01/07/2018"  then	divd(i)= 5572.20	end if	
					
					if rs("memno")=3300 and sldate="01/09/2018"  then	divd(i)= 2873.52	end if	
					if rs("memno")=3300 and sldate="01/08/2018"  then	divd(i)= 2573.52	end if	
					if rs("memno")=3300 and sldate="01/07/2018"  then	divd(i)= 2573.52	end if
					
					if rs("memno")=3341 and sldate="01/11/2018"  then	divd(i)= 420.23		end if	
					if rs("memno")=3341 and sldate="01/10/2018"  then	divd(i)= 420.23		end if	
					if rs("memno")=3341 and sldate="01/09/2018"  then	divd(i)= 420.23		end if	
					if rs("memno")=3341 and sldate="01/08/2018"  then	divd(i)= 420.23		end if	
					if rs("memno")=3341 and sldate="01/07/2018"  then	divd(i)= 420.23		end if
					
					if rs("memno")=3391 and sldate="01/09/2018"  then	divd(i)= 1441.13	end if	
					if rs("memno")=3391 and sldate="01/08/2018"  then	divd(i)= 1441.13	end if	
					if rs("memno")=3391 and sldate="01/07/2018"  then	divd(i)= 1441.13	end if	
					
					if rs("memno")=3770 and sldate="01/10/2018"  then	divd(i)= 2344.53	end if	
					if rs("memno")=3770 and sldate="01/09/2018"  then	divd(i)= 1844.53	end if	
					if rs("memno")=3770 and sldate="01/08/2018"  then	divd(i)= 344.53		end if
					if rs("memno")=3770 and sldate="01/07/2018"  then	divd(i)= 344.53		end if
					
					if rs("memno")=4119 and sldate="01/07/2018"  then	divd(i)= 755.24		end if
					if rs("memno")=4119 and sldate="01/08/2018"  then	divd(i)= 755.24		end if
					
					if rs("memno")=4372 and sldate="01/09/2018"  then	divd(i)= 1783.94	end if	
					if rs("memno")=4372 and sldate="01/08/2018"  then	divd(i)= 1783.94	end if
					if rs("memno")=4372 and sldate="01/07/2018"  then	divd(i)= 1783.94	end if
					
					if rs("memno")=4388 and sldate="01/12/2018"  then	divd(i)= 9728.39 	end if					
					if rs("memno")=4388 and sldate="01/11/2018"  then	divd(i)= 8228.39	end if	
					if rs("memno")=4388 and sldate="01/10/2018"  then	divd(i)= 6728.39	end if	
					if rs("memno")=4388 and sldate="01/09/2018"  then	divd(i)= 6728.39	end if	
					if rs("memno")=4388 and sldate="01/08/2018"  then	divd(i)= 6728.39	end if	
					if rs("memno")=4388 and sldate="01/07/2018"  then	divd(i)= 6728.39	end if
					
					if rs("memno")=4442 and sldate="01/07/2018"  then	divd(i)= 4180.24	end if
					
					if rs("memno")=4508 and sldate="01/12/2018"  then	divd(i)= 10371.80	end if
					if rs("memno")=4508 and sldate="01/11/2018"  then	divd(i)= 10371.80	end if
					if rs("memno")=4508 and sldate="01/10/2018"  then	divd(i)= 9871.80	end if
					if rs("memno")=4508 and sldate="01/09/2018"  then	divd(i)= 9871.80	end if
					if rs("memno")=4508 and sldate="01/08/2018"  then	divd(i)= 9871.80	end if	
					if rs("memno")=4508 and sldate="01/07/2018"  then	divd(i)= 9871.80	end if
					
					if rs("memno")=4515 and sldate="01/01/2019"  then	divd(i)= 3816.55	end if
				    if rs("memno")=4515 and sldate="01/12/2018"  then	divd(i)=  3816.55	end if
					if rs("memno")=4515 and sldate="01/11/2018"  then	divd(i)=  3816.55	end if
					if rs("memno")=4515 and sldate="01/10/2018"  then	divd(i)=  3816.55	end if
					if rs("memno")=4515 and sldate="01/09/2018"  then	divd(i)=  3816.55	end if
					if rs("memno")=4515 and sldate="01/08/2018"  then	divd(i)=  3816.55	end if	
					if rs("memno")=4515 and sldate="01/07/2018"  then	divd(i)= 3800.19	end if	
					if rs("memno")=4556 and sldate="01/07/2018"  then	divd(i)= 369.95		end if
					
					if rs("memno")=4629 and sldate="01/08/2018"  then	divd(i)= 559.00		end if	
					if rs("memno")=4629 and sldate="01/07/2018"  then	divd(i)= 3800.19	end if	
					
					if rs("memno")=2527 and sldate="01/07/2020"  then	divd(i)= 61660.74	end if
					
					'**** 20***
					if rs("memno")=20 and sldate="01/07/2020"  then	divd(i)= 71479.78	end if
					if rs("memno")=20 and sldate="01/08/2020"  then	divd(i)= 71779.78	end if
					if rs("memno")=20 and sldate="01/09/2020"  then	divd(i)= 72079.78	end if				
					if rs("memno")=20 and sldate="01/10/2020"  then	divd(i)= 72379.78	end if					
					if rs("memno")=20 and sldate="01/11/2020"  then	divd(i)= 72679.78	end if					
					if rs("memno")=20 and sldate="01/12/2020"  then	divd(i)= 72979.78 	end if					
					if rs("memno")=20 and sldate="01/01/2021"  then	divd(i)= 73279.78	end if					
					if rs("memno")=20 and sldate="01/02/2021"  then	divd(i)= 73579.78   end if					
					if rs("memno")=20 and sldate="01/03/2021"  then	divd(i)= 73879.78   end if				
					if rs("memno")=20 and sldate="01/04/2021"  then	divd(i)= 74179.78	end if
					if rs("memno")=20 and sldate="01/05/2021"  then	divd(i)= 74479.78	end if	
					if rs("memno")=20 and sldate="01/06/2021"  then	divd(i)= 75078.78	end if						

					
					'**** 2363  ****
					if rs("memno")=2363 and sldate="01/07/2020"  then	divd(i)= 193511.67	end if
					if rs("memno")=2363 and sldate="01/08/2020"  then	divd(i)= 193511.67	end if
					if rs("memno")=2363 and sldate="01/09/2020"  then	divd(i)= 193511.67	end if				
					if rs("memno")=2363 and sldate="01/10/2020"  then	divd(i)= 193511.67	end if					
					if rs("memno")=2363 and sldate="01/11/2020"  then	divd(i)= 193511.67	end if					
					if rs("memno")=2363 and sldate="01/12/2020"  then	divd(i)= 193511.67	end if					
					if rs("memno")=2363 and sldate="01/01/2021"  then	divd(i)= 193511.67	end if					
					if rs("memno")=2363 and sldate="01/02/2021"  then	divd(i)= 193511.67  end if					
					if rs("memno")=2363 and sldate="01/03/2021"  then	divd(i)= 193511.67  end if				
					if rs("memno")=2363 and sldate="01/04/2021"  then	divd(i)= 193511.67	end if
					if rs("memno")=2363 and sldate="01/05/2021"  then	divd(i)= 193511.67	end if	

					'**** 3179  ****
					if rs("memno")=3179 and sldate="01/07/2020"  then	divd(i)= 60665.11	end if
					if rs("memno")=3179 and sldate="01/08/2020"  then	divd(i)= 61865.11	end if
					if rs("memno")=3179 and sldate="01/09/2020"  then	divd(i)= 63065.11	end if				
					if rs("memno")=3179 and sldate="01/10/2020"  then	divd(i)= 64265.11	end if					
					if rs("memno")=3179 and sldate="01/11/2020"  then	divd(i)= 65465.11	end if					
					if rs("memno")=3179 and sldate="01/12/2020"  then	divd(i)= 66665.11	end if					
					if rs("memno")=3179 and sldate="01/01/2021"  then	divd(i)= 67865.11	end if					
					if rs("memno")=3179 and sldate="01/02/2021"  then	divd(i)= 69065.11   end if					
					if rs("memno")=3179 and sldate="01/03/2021"  then	divd(i)= 70265.11   end if				
					if rs("memno")=3179 and sldate="01/04/2021"  then	divd(i)= 71465.11	end if
					if rs("memno")=3179 and sldate="01/05/2021"  then	divd(i)= 72665.11	end if					

					'****** 4986  *******
					if rs("memno")=4986 and sldate="01/07/2020"  then	divd(i)= 2500		end if
					if rs("memno")=4986 and sldate="01/08/2020"  then	divd(i)= 2505.83	end if
					if rs("memno")=4986 and sldate="01/09/2020"  then	divd(i)= 2505.83	end if
					if rs("memno")=4986 and sldate="01/10/2020"  then	divd(i)= 2505.83	end if					
					if rs("memno")=4986 and sldate="01/11/2020"  then	divd(i)= 2505.83	end if					
					if rs("memno")=4986 and sldate="01/12/2020"  then	divd(i)= 2505.83	end if					
					if rs("memno")=4986 and sldate="01/01/2021"  then	divd(i)= 2505.83	end if					
					if rs("memno")=4986 and sldate="01/02/2021"  then	divd(i)= 2505.83	end if					
					if rs("memno")=4986 and sldate="01/03/2021"  then	divd(i)= 2505.83	end if				
					if rs("memno")=4986 and sldate="01/04/2021"  then	divd(i)= 2505.83	end if
					if rs("memno")=4986 and sldate="01/05/2021"  then	divd(i)= 2505.83	end if
					
					

					'***********************************************************************			   
					desc   = "　"&formatNumber(divd(i))&" X "&rate&"/"&"100"&"/12"
				else
					desc   = "　0 X"&rate&"/"&"100"&"/12"
				end if
		   
           xamt   =formatnumber( divd(i)*rate/100/12,2)*1
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
</td>
</tr>
</table>

<BR>
	<font face="arial, helvetica, sans-serif" size="3" ><b>　　　　</b></font><br>
	<font face="arial, helvetica, sans-serif" size="5" ><b>　　　　承繼人 1. <%=B1%> ,   承繼人 2. <%=B2%> </b></font><br>
	<font face="arial, helvetica, sans-serif" size="3" ><b>　　　　</b></font><br>
	<font face="arial, helvetica, sans-serif" size="3" ><b>　　　　</b></font><br>

<font face="arial, helvetica, sans-serif" size="3" ><b>　　　　</b></font><br>

<font face="arial, helvetica, sans-serif" size="3" ><b>　　　　註 - (1) 本年度的協會費將由本社支付 </b></font><br>
<%if ttlamt > 100 then %>
<%if mstatus ="M" or mstatus="T" then %>
<font face="arial, helvetica, sans-serif" size="3" ><b>　　　　　 - (2) 閣下於本年度所賺取的股息為<%=formatnumber(ttlamt,2)%>元，以支票支付。</b></font><br>     
<%else%>
<font face="arial, helvetica, sans-serif" size="3" ><b>　　　　　 - (2) 閣下於本年度所賺取的股息為<%=formatnumber(ttlamt,2)%>元，於<%=pyear%>年<%=int(mon)%>月<%=nday%>日自動轉賬入閣下銀行戶口。</b></font><br>
<%end if%>
<%else%>
<font face="arial, helvetica, sans-serif" size="3" ><b>　　　　　 - (2) 閣下於本年度股息為<%=formatnumber(ttlamt,2)%>元，於<%=pyear%>年<%=int(mon)%>月<%=nday%>日存入閣下的股金內。</b></font><br>
<%end if%>

<font face="arial, helvetica, sans-serif" size="3" ><b>　　　　　 - (3) 如閣下想接收本社最新最快的消息，請立刻安裝WhatsApp程式，並將電話號碼 5476 6906 </b></font><br>
<font face="arial, helvetica, sans-serif" size="3" ><b>　　　　　  　　加入閣下的電話簿內即可。同時亦可隨時登入本社網址 www.wsdscu.org 瀏覽所有通告及最新消息。</b></font><br>
<font face="arial, helvetica, sans-serif" size="3" ><b>　　　　　 - (4) 為響應環保，本社將不再郵寄活動通告給退休社員，個別退休社員如有需要，可聯絡本社職員。</b></font><br>
<BR>
<BR>
<BR>


<%
       for i = 1 to 12
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
