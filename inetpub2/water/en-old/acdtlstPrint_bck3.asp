<!-- #include file="../conn.asp" -->

<%
id  =request.form("memno")

if id="" then
    response.redirect "acdetaillst.asp"
end if

chkdate = request.form("stdate")
enddate = request.Form("enddate")

if id="" or chkdate="" then
    response.redirect "acdetaillst.asp"
end if

mndate=right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
xmon = mid(chkdate,4,2)
xyr1  =cint( right(chkdate,4)) 
dd    = cint(left(chkdate,2))

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
ndate =xyr1&right("0"&xmn,2)&right("0"&xdd,2)
enddateS= cint(right(enddate,4)) & mid(enddate,4,2) & cint(left(enddate,2))
enddateL= cint(right(enddate,4)) &"/" & mid(enddate,4,2) &"/"& cint(left(enddate,2))
'response.Write(enddate)
'response.Write("<BR>")
'response.Write(enddateS)
'response.Write("<BR>")
'response.Write(enddateL)

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


'Loan
       
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

scode(1) = "�Ѫ����l"

if xlnnum <> "" then
    'SQl = "select  *,convert(char(10),ldate,102) as Expr1  from loan  where memno='"&id&"' and ldate  >='"&xldate&"' order by memno,ldate,right(code,1),left(code,1) "
    'SQl = "select  *,convert(char(10),ldate,102) as Expr1  from loan  where memno='"&id&"' and ldate  >='"&xldate&"' and ldate <= '"&enddateL&"' order by memno,ldate,right(code,1),left(code,1) "
    SQl = "select  *,convert(char(10),ldate,102) as Expr1  from loan  where memno='"&id&"' and ldate <= '"&enddateL&"' order by memno,ldate,right(code,1),left(code,1) "

    Set rs = Server.CreateObject("ADODB.Recordset")	
    rs.open sql, conn,2  

    cc = 0
    xx = 1
    qx = 0 = 0
    MX=0
    zero = 0
    count = 0

    do while not rs.eof
      lncode(xx) = rs("code")
      xdate = right("0"&day(rs("ldate")),2)&"/"&right("0"&month(rs("ldate")),2)&"/"&year(rs("ldate"))
      
      xyear = year(rs("ldate"))
      xmon  = month(rs("ldate"))
      xday  = day(rs("ldate"))
      xxdate = xyear&right("0"&xmon,2)&right("0"&xday,2)
      if xxdate  > ndate and xxdate < enddateS then
           
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
                      CASE "ME"
                           lcode(xx) ="�Ȧ���" 
                      CASE "MF"
                           lcode(xx) ="�Q�����" 
                      CASE "NE"
                           lcode(xx) ="�w�в��"            
                      CASE "D1"
             
                      CASE "D0"
                          if rs("amount") > 0 then
                            lcode(xx) ="�U�ڲM��"
                          end if   
                      case "ME","NE"
                          mx = 0  
                 end select 
               
                 IF (RS("CODE")="F3" or rs("code")="F1" or rs("code")="F2") AND MX= 0 THEN
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
                                IF lnflag = "Y" then 
                                    zero = 0
                                    lniamt(xx) =  rs1(0) 
                                    lcode(xx) ="+ �s�U  ="
                                    lnbal(xx) = rs1(2)  
                                    bal = lnbal(xx)          
                                ELSE
                                    lnbal(xx) = rs("amount")
                                    lniamt(xx) =""
                                    bal = lnbal(xx)
                                    lcode(xx) ="�s�U"  
                                END IF                     
                            end if
                            rs1.close 
                            MX = 0     
                        end if 
                    end if
                END IF

                xx = xx + 1
                bb = xx 
                if rs("code")="D0" and rs("amount")= 0 then
                   xx = xx -  1
                zero= 1
                end if
           end if 
           
           if left(rs("code"),1)="E" then
              mx = mx + 1
           end if
           
           if left(rs("code"),1) ="F" and xdate=lndate(xx-1) and mx = 1 then
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
           
       else        
            select case rs("code")           
                  case "0D","D1" 
                         bal = rs("amount")
                  case  "E0","E1" , "E4" , "E2" , "E3" , "E6" 
                        bal = bal - rs("amount")
                  case  "ER"
                       bal = bal + rs("amount")
                  case  "ME" 
    
            end select  
        end if 

       rs.movenext 
    loop
    rs.close
end if


'share
sql1 = "select  * from share  where memno= '"&id&"'    order by memno,ldate,code  "
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
   
    'if xxdate  > ndate then 
    if xxdate  > ndate and xxdate < enddateS then
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
            case "0A","A1","A2","A3","C0","C1","C3"   
                    xbal = xbal + rs1("amount")
        end select
        sbal(xx) = xbal
    
        select case rs1("code")
              case "A1"
                   scode(xx) = "�Ȧ���b"
              case "A2"
	            scode(xx) ="�w����b"
              case "A3"
	            scode(xx) ="�{���s��"
              case "B0"
                   scode(xx)="�Ѫ��ٴ�"
              case "B1"
                   if rs1("amount") >0  then  
                       scode(xx)="�h��"
                    else
                        scode(xx)="�h�ٶU��"
                   end if
              CASE "AI"
                    scode(xx) ="���" 
              CASE "D1"
                   scode(xx) ="�s�U"  
              CASE "B0"
                    scode(xx) ="�{���h��"
              case "B3"
                    scode(xx) ="�h�ٲ{��"                
              case "C0"
                   scode(xx)="�Ѯ�"         
              case "C1"
                   scode(xx)="�Ѯ��Ȧ��b" 
              case "C3"
                    scode(xx)="�Ѯ��{����b" 
              case "G0","G1","G2","G3"
                    scode(xx) = "��|�O" 
              case "H0","H1","H2","H3"
                    scode(xx) = "�J�|�O"
        end select
        
        sdate(xx) = ssdate
        if rs1("amount") > 0 then       
            samt(xx) =rs1("amount")
        end if
        xx = xx + 1
        aa = xx
    else
        if  left(rs1("code"),1)="G" or left(rs1("code"),1)="H" or left(rs1("code"),1)="B" then        
            xbal=xbal-rs1("amount")
        else       
            xbal = xbal +rs1("amount")
        end if        
    end if
 
   
   rs1.movenext
loop
rs1.close
xx = 1
bb = bb - countrs
if aa > bb then
   smax = aa
else
   smax = bb
end if


'output format
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
	objFile.Write "���ȸp���u�x�W���U��"
	objFile.WriteLine ""
	objFile.Write "�ӤH��f�d�ߦC��"
	objFile.WriteLine ""
 	objFile.Write "�����W�� : "&memname&memcname
	objFile.WriteLine ""
 	objFile.Write "�����s�� : "&id
	objFile.WriteLine ""
    objFile.WriteLine ""
    if guid1 <>"" then 
        objFile.Write  "1.��O�H"
        objFile.WriteLine ""
        objFile.Write  "�����s�� : "&guid1&" �����s�� : "&guname1
        objFile.WriteLine ""
        objFile.Write "�x�W���l : "&formatnumber(gusave1,2)
        objFile.WriteLine ""
      
      
    end if
    if guid2 <>"" then
          
        objFile.Write  "2.��O�H"
        objFile.WriteLine ""
        objFile.Write  "�����s�� : "&guid2&" �����s�� : "&guname2
        objFile.WriteLine ""
        objFile.Write "�x�W���l : "&formatnumber(gusave2,2)
        objFile.WriteLine ""
      
       
    end if
    if guid3 <>"" then 
          
        objFile.Write  "3.��O�H"
        objFile.WriteLine ""
        objFile.Write  "�����s�� : "&guid3&" �����s�� : "&guname3
        objFile.WriteLine ""
        objFile.Write "�x�W���l : "&formatnumber(gusave3,3)
        objFile.WriteLine ""
      
      
    end if
    if guoid1 <>"" then 
                
        objFile.Write  "1.��O��L�H"
        objFile.WriteLine ""
        objFile.Write  "�����s�� : "&guiod1&" �����s�� : "&guoname1
        objFile.WriteLine ""
        objFile.Write "�x�W���l : "&formatnumber(guosave1,2)
        objFile.WriteLine ""
      
      
    end if
    if guoid2 <>"" then 
          
        objFile.Write  "2.��O��L�H"
        objFile.WriteLine ""
        objFile.Write  "�����s�� : "&guiod2&" �����s�� : "&guoname2
        objFile.WriteLine ""
        objFile.Write "�x�W���l : "&formatnumber(guosave2,2)
        objFile.WriteLine ""
      
       
    end if
    if guoid3 <>"" then 
          
       objFile.Write  "3.��O��L�H"
        objFile.WriteLine ""
        objFile.Write  "�����s�� : "&guiod3&" �����s�� : "&guoname3
        objFile.WriteLine ""
        objFile.Write "�x�W���l : "&formatnumber(guosave3,2)
        objFile.WriteLine ""
      
      
    end if
	objFile.Write left("  ���"&spaces,10)
	objFile.Write left("    �Ѫ�"&spaces,13)
	objFile.Write right("   ���l",16)
	objFile.Write left("     ���O"&spaces,15)
	objFile.Write left("  ���"&spaces,8)
        objFile.Write left(" �U�ڽs��"&spaces,8)
	objFile.Write left(" �Q��"&spaces,8)
        objFile.Write left("�C���ٴ�"&spaces,8)
	objFile.Write left("�s�U�`�B/���l"&spaces,10)
	objFile.Write left("���O"&spaces,6)
	objFile.WriteLine ""
	for idx = 1 to 120
		objFile.Write "-"
	next
	
	objFile.WriteLine ""
	xx = 1
    do while xx <= smax
        if sdate(xx)<>"" or lnnum(xx)<>"" then
            if sbal(xx) > 0 then
                if sdate(xx)<>"" then 
                    objFile.Write left(" "&sdate(xx)&spaces,12) 
                else
                    objFile.Write right(spaces,12)
                end if
                
                if samt(xx) <>"" then 
                 objFile.Write right(spaces&formatNumber(samt(xx),2),10)  
                else
                    objFile.Write right(spaces,10)
                end if
                
                if sbal(xx) <>"" then 
                 objFile.Write right(spaces&formatNumber(sbal(xx),2),13)  
                else
                    objFile.Write left(spaces,13)
                end if
                
                if scode(xx)<>"" then                                               
                 objFile.Write left("  "&scode(xx)&spaces,8)
                else
                  objFile.Write right(spaces,12)
                end if
            else
                objFile.Write right(spaces,47)
            end if
            objFile.Write  "|"

            if lnnum(xx)<>"" then
                objFile.Write right(spaces&lndate(xx),11)             
                objFile.Write left("   "&lnnum(xx)&spaces,11) 
                if lniamt(xx)<>"" then 
                    objFile.Write right(spaces&formatNumber(lniamt(xx),2),10)
                else
                    objFile.Write left(spaces,10)
                end if

                if lnramt(xx)<>"" then            
                    objFile.Write  right(spaces&formatNumber(lnramt(xx),2),10)
                else
                    objFile.Write  left(spaces,10)
                end if                                 

                if lnbal(xx)<>"" then 
                    objFile.Write  right(spaces&formatNumber(lnbal(xx),2),15)
                else
                    objFile.Write right(spaces,15)
                end if
                objFile.Write left("    "&lcode(xx)&spaces,12)
            end if
            objFile.WriteLine ""
        else
            xx = 501
        end if
        xx = xx + 1
    loop            

 
	for idx = 1 to 120
		objFile.Write "-"
	next

	objFile.Close
	set rs=nothing
    set rs1=nothing
	conn.close
	set conn=nothing
	response.redirect "../txt/"&session("username")&".txt"
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
<table border="0" cellpadding="0" cellspacing="0">
	<tr height="30" valign="top" align="center">
		<td colspan="15"><font size="4">���ȸp���u�x�W���U��</font></td>		
        <td width="200" align="right">��� : <%=mndate%></td>  
	</tr>
	<tr height="35" valign="top" align="center">
		<td colspan="15"><font size="4">�ӤH��f�d�ߦC��</font></td>
	</tr>
	<tr height="35" valign="top" align="center">
		<td colspan="15"><font size="4">�����W��<%=memname%><%=memcname%> </font></td>
	</tr>
	<tr height="35" valign="top" align="center">
		<td colspan="15"><font size="4">�����s��<%=id%></font></td>
	</tr>
</table>


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

<table border="0" cellspacing="2" cellpadding="1" align="center" >
	<tr valign="top" align="center">
		
		<td><font size="2" >���</font></td>
               
		<td><font size="2" >�Ѫ�</font></td>
               
		<td><font size="2" >���l</font></td>
                            
		<td><font size="2">���O</font></td>
               
               <td bgcolor="#FFFFF"> </td>
		<td><font size="2" >���</font></td>
               
		<td><font size="2" >�U�ڽs��</font></td>
                
		<td><font size="2" >�Q��</font></td>
               
		<td><font size="2" >�C���ٴ�<font></td>
                
		<td><font size="2" >�s�U�`�B/���l</font></td>
              
		<td><font size="2" >���O</font></td>            
	</tr>
	<tr><td colspan=11><hr></td></tr>
	
<%
xx = 1
do while xx <= smax
    if sdate(xx)<>"" or lnnum(xx)<>"" then
    %>
    <tr>
        <%if sbal(xx) > 0   then %>   	   
  		    <td><font size="1"><%=sdate(xx)%></font></td>
                   
            <%if samt(xx) <> ""   then %>
                <td align="right"><font size="1"><%=formatNumber(samt(xx),2)%></font></td>
            <%else%>
                <td></td>
            <%end if %>  

                         
            <%if sbal(xx) <> ""   then %>
                <td align="right"><font size="1"><%=formatNumber(sbal(xx),2)%></font></td>
            <%else%>
                <td></td>
            <%end if %>     
     

     		
 		    <td><font size="1"><center><%=scode(xx)%></center></font></td>
        <%else%>
                <td></td>
                <td></td>
                <td></td>         
                <td></td>                                
        <%end if%> 
        <td bgcolor="red"> </td>	
        
        <%if lnnum(xx)<>"" then %>   
            
            <td><font size="1"><%=Lndate(xx)%></font></td>	
            <td align="right"><font size="1"><%=lnnum(xx)%></font></td>
                   	
            <%if lniamt(xx)<>"" then %>
                <td align="right"><font size="2"><%=formatNumber(lniamt(xx),2)%></font></td>
            <%else%>
                <td align="right"><%=lniamt(xx)%></td>
            <%end if%>

            <%if lnramt(xx) <>""   then %>           
                <td align="right"><font size="2"><%=formatNumber(lnramt(xx),2)%></font></td>
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
                
            <%  
                end if
            end if                          
            %> 
                          	
            <%if lnbal(xx)<>"" then %>                
                <td align="right"><font size="1"><%=formatNumber(lnbal(xx),2)%></font></td>
            <%else%>
                <td><%=lnbal(xx)%></td>
            <%end if%>

            <td><font size="1"><center><%=lcode(xx)%></center></font></td>
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


