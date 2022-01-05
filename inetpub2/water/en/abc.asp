<!-- #include file="../conn.asp" -->

<%
xdate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())

mPeriod = request.form("mPeriod")
mperiod = "200701"
yy  = cint(left(mperiod,4))
mm = cint(right(mPeriod,2))
xyy =  yy
xmm = mm + 1
if xmm= 13 then
   xyy = xyy + 1
   xmm = xmm - 12
end if
xdd= 1

nPeriod = left(mPeriod,4)&"."&right(mPeriod,2)&".01"
pPeriod = xyy&"."&right("0"&xmm,2)&".01"
xperiod = xyy&"/"&right("0"&xmm,2)&"/"&"01"


ryy = yy
rmm = mm - 1
if rmm=0 then
    ryy = yy - 1
    rmm = 12
end if
RPeriod = ryy&rmm

server.scripttimeout = 1800
         set rs = conn.execute("select memno,code,amount from loan where ldate>='2007/01/01' and  left(code,1)='D' order by memno,ldate,code ")
         memno=rs("memno")

         do while not rs.eof
            if memno<>rs("memno") then
               actttl = a - b
               gttl = gttl + actttl
               response.write(memno)
               response.write("**")
               response.write(actttl)
               response.write("==")

               a = 0
               b = 0
               memno = rs("memno")
            end if
            select case rs("code")
                   case "D1"
                       a = rs("amount")
                   case "D0"
                       b = rs("amount")
           end select
           rs.movenext
           loop
           rs.close
           response.write("()")
           response.write(gttl)
           response.end              
       
         set rs = conn.execute("select count(*) from memmaster where memdate>='"&nperiod&"' and memdate <'"&xperiod&"' ")
         if not rs.eof then
            nmemcnt = rs(0)
         end if 
         rs.close
 
         set rs = conn.execute("select count(*) from memmaster where wdate >='"&nperiod&"' and wdate <'"&xperiod&"' ")
         if not rs.eof then
            omemcnt = rs(0)
         end if 
         rs.close


         memcnt = 0
         set rs = server.createobject("ADODB.Recordset")   
         savettl = 0
         sql = "select a.memno,a.code,sum(a.amount) as ammt  from share a,memmaster b where a.memno=b.memno and a.ldate<'"&nperiod&"' and b.memdate<'"&nperiod&"' and (b.wdate is null or b.wdate>='"&nperiod&"')  group by a.memno,a.code  order by a.memno,a.code "
         rs.open sql, conn, 1, 1 
         memno= rs("memno")
         do while  not rs.eof 
            if memno <> rs("memno") then 
               if savettl > 0 then
                  memcnt = memcnt + 1
               end if
               savettl = 0
               memno= rs("memno")
            end if
               select case rs("code")
                      case "0A","A1","A2","A3","C0","C1","C3","A0"
                           savettl = savettl + rs("ammt")
                      case "B0","B1","G0","G1","G3","H0","H1","H3"
                           savettl = savettl - rs("ammt") 
               end select 
 
         rs.movenext
         loop
              if savettl > 0 then
                  memcnt = memcnt + 1
               end if
         rs.close  



         set rs = conn.execute("select count(*) from loanrec where lndate < '"&nperiod&"' and cleardate is null    ")
         if not rs.eof    then
            ttllncnt = rs(0)
         END IF 

         rs.close
         set rs = conn.execute("select count(*) from loanrec where lndate < '"&nperiod&"' and cleardate>'"&nperiod&"'  ")
         if not rs.eof    then
             ttllncnt =  ttllncnt+ rs(0)
         END IF 
         rs.close
         set rs = conn.execute("select count(*) from loanrec where lndate >= '"&nperiod&"' and lndate <'"&xperiod&"'   ")
         if not rs.eof    then
             nwlncnt =  rs(0)
         END IF 
         rs.close

         set rs = conn.execute("select appamt,bal,convert(char(10),lndate,102) as slndate   from loanrec where repaystat='N' and lndate< '"&xperiod&"'  ")
         do while not rs.eof
          
            if rs("slndate")>= nperiod  then   
               ttllnamt = ttllnamt + rs("appamt")
               ttlbal   = ttlbal + rs("bal")
           else
               ottllnamt = ottllnamt + rs("appamt")
               ottlbal = ittlbal + rs("bal")
           end if
           lncount  = lncount + 1 
           rs.movenext
        loop
        rs.close
       clncnt = 0
        ttlnwlnamt = 0
         set rs = conn.execute("select code,convert(char(10),ldate,102) as pydate ,amount from loan where  ldate <'"&xperiod&"' order by memno,ldate,code ")
         do while not rs.eof
         

   
              if rs("pydate") >=nperiod    then 
               select case rs("code")
                      case  "E1"
                           lbnkamt = lbnkamt + rs("amount")
                      case "E2"
                          lsadamt = lsadamt + rs("amount")
                      case  "E3" 
                          lchamt = lchamt + rs("amount")
                      case "E0"
                          ajlnamt = ajlnamt + rs("amount") 
                      case  "F1"
                           ibnkamt = ibnkamt + rs("amount")
                      case "F2"
                          isadamt = isadamt + rs("amount")
                      case  "F3" 
                          ichamt = ichamt + rs("amount")
                      case "F0"
                          ajintamt = ajintamt + rs("amount")
                      case "D0"
                           if rs("amount") > 0 then
                              clncnt = clncnt + 1
                           end if
                           clnamt = clnamt + rs("amount")
                           
                      case "D1"
                          
                           nwlnamt = nwlnamt + rs("amount") 
                      case "ET"
                           esavamt = esavamt + rs("amount")
                      case "FT"
                           fsavamt = fsavamt + rs("amount")
             end select                    
             else
                 select case rs("code")

                      case "D0","E1","E0","E2","E3"
                           
                          ttlnwlnamt = ttlnwlnamt - rs("amount")
                      case "D1" ,"0D"  
                                             
                           ttlnwlnamt = ttlnwlnamt + rs("amount")
                  end select
             end if
          
             rs.movenext
             loop
         rs.close          



         set rs1 = conn.execute("select code,convert(char(10),ldate,102) as pydate  ,amount from share where ldate<'"&xperiod&"'order by memno,ldate,code ")
         do while not rs1.eof
            curdate = rs1(1)

            if rs1("pydate")<  nperiod then
               select case rs1("code")
                      case "0A", "A1","A2","A3","C0","C1","C3","A0"
                           ttlamt = ttlamt + rs1("amount")
                      case "B0","B1","B2","B3","G0","G1","G3","H0","H1","H3"
                           ttlamt = ttlamt - rs1("amount")
             end select
             else
              
               select case rs1("code")
                      case  "A1"
                           bnkamt = bnkamt + rs1("amount")
                      case "A2"
                           sadamt = sadamt + rs1("amount")
                      case  "A3" ,"0A"
                          chamt = chamt + rs1("amount")
                      case "A0"
                           ajshamt = ajshamt + rs1("amount")
                      case "C3"
                           divamt3 = divamt3 + rs1("amount")
                      
                      case "C1"
                           divamt1 = divamt1 + rs1("amount")
                      case "C0"
                           ajdivamt = ajdivamt + rs1("amount")
                      case "B1"
                           withdamt = withdamt + rs1("amount")
                      case "B0"
                            ajwdamt = ajwdamt + rs1("amount")
                      
                  case "G3"
			Gamt3 = Gamt3+rs1("amount")
                  case "H3"
			Hamt3 = Hamt3+rs1("amount")
                 case "G2"
			Gamt2 = Gamt2+rs("amount")
                  case "H2"
			Hamt2 = Hamt2+rs1("amount")
                 case "G1"
			Gamt1 = Gamt1+rs1("amount")
                  case "H1"
			Hamt1 = Hamt1+rs1("amount")
             end select                    
             end if
      
             rs1.movenext
             loop
         rs1.close
         
         csttlamt = bnkamt+sadamt+chamt-gamt3-hamt3+ ajshamt
         cwttlamt = withdamt+ajwdamt
         cajttlamt = ajshamt+ ajdivamt+ajlnamt+ajintamt
         cloanamt   = nwlnamt - clnamt
         ttldiv = divamt1  + divamt3+ajdivamt
         ttlgamt = Gamt1 + Gamt2 + Gamt3
         ttlhamt = Hamt1 + Hamt2 + Hamt3
         cpayamt  = lbnkamt +lsadamt+lchamt+ ajlnamt
         cintamt  = ibnkamt +isadamt+ichamt+ajintamt
         ttlbnk =  lbnkamt+bnkamt + ibnkamt 
         ttlsad =  lsadamt+sadamt + isadamt
         ttlch  =  lchamt +chamt  + ichamt + divamt3 
         ttlrec = ttlbnk + ttlsad + ttlch + cajttlamt 
         gttlamt = csttlamt+ttlamt-cwttlamt+ttldiv
         actln  = nwlnamt - clnamt 
         payamt = actln + withdamt 
	 ttlpay = payamt + ajwdamt 
         glnamt = ttlnwlnamt 
         actlnamt = glnamt + cloanamt - cpayamt 
         actlncnt = ttllncnt +nwlncnt - clncnt
         ttlmem = memcnt 
         actmem = memcnt + nmemcnt -omemcnt
         gttlrate= round(actlnamt / gttlamt*100,0)
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
	objFile.Write  left(space2,20)
	objFile.Write "水務署員工儲蓄互助社"
	objFile.WriteLine ""
	objFile.Write  left(space2,21)
	objFile.Write "每月帳統計列表"
	objFile.WriteLine ""	
	objFile.Write "日期 :"
        objFile.Write  xdate
	objFile.WriteLine ""	
	objFile.Write left(mPeriod&spaces,10)
	objFile.Write left("    銀行"&spaces,12)
	objFile.Write left("    庫房"&spaces,12)
	objFile.Write left("  現金"&spaces,12)
        objFile.Write left("  調整"&spaces,12)
	objFile.Write left("  "&mm&" 月"&spaces,15)   
 	objFile.Write left("  "&rmm&" 月"&spaces,12)
	objFile.Write left(" 總  結"&spaces,10)    
	objFile.WriteLine ""
	for idx = 1 to 107
		objFile.Write "-"
	next
        
	objFile.WriteLine ""   


	objFile.Write left("  股金　"&spaces,10)
	objFile.Write right(spaces&formatnumber(bnkamt,2),12)
	objFile.Write right(spaces&formatnumber(sadamt,2),12)
	objFile.Write right(spaces&formatnumber(chamt-Gamt3-Hamt3 ,2),12)
        objFile.Write right(spaces&formatNUMBER(ajshamt,2),12)
	objFile.Write right(spaces&formatnumber(csttlamt,2),15)
	objFile.Write right(spaces&formatnumber(ttlamt,2),15)
	objFile.Write right(spaces&formatnumber(gttlamt,2),15)
	objFile.WriteLine ""
	objFile.Write left("  退股　"&spaces,10)
	objFile.Write right(spaces&formatnumber(withdamt,2),12)
        objfile.Write left(spaces,24)
        objFile.Write right(spaces&formatNumber(ajwdamt,2),12)
	objFile.Write right(spaces&formatnumber(cwttlamt,2),15)
       
	objFile.WriteLine ""
	objFile.Write left("  股息　"&spaces,10)
	objFile.Write right(spaces&formatnumber(divamt1,2),12)
	objFile.Write right(spaces&formatnumber(divamt2,2),12)
	objFile.Write right(spaces&formatnumber(divamt3,2),12)
        objFile.Write right(spaces&formatNumber(ajdivamt,2),12)
	objFile.Write right(spaces&formatnumber(ttldiv,2),15)
	objFile.WriteLine ""
	objFile.Write left("  貸款　"&spaces,10)
	objFile.Write right(spaces&formatnumber(actln,2),12)
        objfile.Write left(spaces,24)
        objFile.Write left(spaces,12)
	objFile.Write right(spaces&formatnumber(cloanamt,2),15)
	objFile.Write right(spaces&formatnumber(glnamt,2),15)
        objFile.Write right(spaces&formatnumber(actlnamt,2),15)
	objFile.WriteLine ""
	objFile.Write left("  還款　"&spaces,10)
	objFile.Write right(spaces&formatnumber(lbnkamt,2),12)
	objFile.Write right(spaces&formatnumber(lsadamt,2),12)
	objFile.Write right(spaces&formatnumber(lchamt,2),12)
        objFile.Write right(spaces&formatnumber(ajlnamt,2),12)
	objFile.Write right(spaces&formatnumber(cpayamt,2),15)
	objFile.WriteLine ""
	objFile.Write left("  利息　"&spaces,10)
	objFile.Write right(spaces&formatnumber(ibnkamt,2),12)
	objFile.Write right(spaces&formatnumber(isadamt,2),12)
	objFile.Write right(spaces&formatnumber(ichamt,2),12)
        objFile.Write right(spaces&formatnumber(ajintamt,2),12)
	objFile.Write right(spaces&formatnumber(cintamt,2),15)
	objFile.WriteLine ""
	objFile.Write left("  入會費"&spaces,10)
	objFile.Write right(spaces&formatnumber(gamt1,2),11)
	objFile.Write right(spaces&formatnumber(gamt2,2),12)
	objFile.Write right(spaces&formatnumber(gamt3,2),12)
        objFile.Write left(spaces,12)
	objFile.Write right(spaces&formatnumber(ttlgamt,2),15)
	objFile.WriteLine ""
	objFile.Write left("  協會費"&spaces,10)
	objFile.Write right(spaces&formatnumber(hamt1,2),11)
	objFile.Write right(spaces&formatnumber(hamt2,2),12)
	objFile.Write right(spaces&formatnumber(hamt3,2),12)
        objFile.Write left(spaces,12)
	objFile.Write right(spaces&formatnumber(ttlHamt,2),15)
        objFile.Write left(spaces,15)
        objFile.Write left("       貸款/股金"&spaces,15) 
	objFile.WriteLine ""
	for idx = 1 to 107
		objFile.Write "-"
	next
	objFile.WriteLine "" 
	objFile.Write left("  收　入"&spaces,10)
	objFile.Write right(spaces&formatnumber(ttlbnk,2),12)
	objFile.Write right(spaces&formatnumber(ttlsad,2),12)
	objFile.Write right(spaces&formatnumber(ttlch,2),12)
        objFile.Write right(spaces&formatnumber(cajttlamt,2),12)
	objFile.Write right(spaces&formatnumber(ttlrec,2),15)
        objFile.Write left("      總額"&spaces,15)
	objFile.Write right(spaces&formatnumber(gttlrate,2)&"%    ",15) 
	objFile.WriteLine ""
	objFile.WriteLine "" 
	objFile.Write left("  支　出"&spaces,10)
	objFile.Write right(spaces&formatnumber(payamt,2),12)
 	objfile.Write left(spaces,24)
        objFile.Write right(spaces&formatnumber(ajwdamt,2),12)
	objFile.Write right(spaces&formatnumber(ttlpay,2),15)
	objFile.WriteLine ""
	objFile.WriteLine ""
	objFile.WriteLine ""

	objFile.Write  space(10)&"貸款總數於"&mm&"月"&yy&"年前 : + "  
        objFile.Write right(spaces&formatnumber(ttllncnt,0),5)
	objFile.WriteLine ""
	objFile.Write space(10)&"新貸款總數             : + "
        objFile.Write right(spaces&formatnumber(nwlncnt,0),5)
        objFile.WriteLine ""	
	objFile.Write space(10)&"已清數貸款總數         : - "
        objFile.Write right(spaces&formatnumber( clncnt,0),5)
        objFile.WriteLine ""
	objFile.Write space(10)&"貸款總數合共           : = "
        objFile.Write right(spaces&formatnumber(actlncnt,0),5)
        objFile.WriteLine ""
        objFile.WriteLine ""
	objFile.Write  space(10)&"社員總數於"&mm&"月"&yy&"年前 :  "  
        objFile.Write right(spaces&formatnumber(ttlmem,0),5)
	objFile.WriteLine ""
	objFile.Write  space(10)&"社員退社總數          : - "  
        objFile.Write right(spaces&formatnumber(omemcnt,0),5)
	objFile.WriteLine ""
	objFile.Write  space(10)&"新社員總數            : + "   
        objFile.Write right(spaces&formatnumber(nmemcnt,0),5)
	objFile.WriteLine ""
	objFile.Write  space(10)&"社員總數合共          : = "  
        objFile.Write right(spaces&formatnumber(actmem,0),5)
	objFile.WriteLine ""
	objFile.Close

	
	
	response.redirect "../txt/"&session("username")&".txt"
end if
%>
<html>
<head>
<title>每月帳統計列表</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
<center>
<table border="0" cellpadding="0" cellspacing="0">
<tr><td><font size="2">水務署員工儲蓄互助社</td>
<tr><td>每月帳統計列表<td></tr>
<tr><td>日期 :<%=xdate%></td></font></tr>
</table>
<table border="0" cellspacing="1" cellpadding="4" align="center" bgcolor="336699">
	<tr bgcolor="#330000" align="center">
		<td><font size="2" color="#FFFFFF"><%=yy%><%=mm%></font></td>
		<td><font size="2" color="#FFFFFF">銀　行</font></td>
		<td><font size="2" color="#FFFFFF">庫　房</font></td>
		<td><font size="2" color="#FFFFFF">現　金</font></td>
		<td><font size="2" color="#FFFFFF">調　整</font></td>
		<td><font size="2" color="#FFFFFF"><%=mm%>月</font></td>
		<td><font size="2" color="#FFFFFF"><%=rmm%>月</font></td>
		<td><font size="2" color="#FFFFFF">總 　結</font></td>

	</tr>
        <tr bgcolor="#FFFFFF">
             <td>  股金　</td>
             <td align="right" ><%=formatnumber(bnkamt,2)%></td>
	     <td align="right" ><%=formatnumber(sadamt,2)%></td>
	     <td align="right" ><%=formatnumber(chamt-Gamt3-Hamt3,2)%></td>
             <td align="right" ><%=formatNUMBER(ajshamt,2)%></td>
	     <td align="right" ><%=formatnumber(csttlamt,2)%></td>
	     <td align="right" ><%=formatnumber(ttlamt,2)%></td>
	     <td align="right" ><%=formatnumber(gttlamt,2)%></td>
        </tr> 
         <tr bgcolor="#FFFFFF">
             <td>  退股　</td>
             <td align="right" ><%=formatnumber(withdamt,2)%></td>
	     <td></td>
	     <td></td>
             <td align="right" ><%=formatNUMBER(ajwdamt,2)%></td>	     
	     <td align="right" ><%=formatnumber(cwttlamt,2)%></td>
             <td></td>
	     <td></td>
        </tr>  
         <tr bgcolor="#FFFFFF">
             <td>  股息　</td>
             <%if divamt1 <> 0 then %>
             <td align="right" ><%=formatnumber(divamt1,2)%></td>
             <%else%>
             <td></td>
             <%end if%>
             <%if divamt2 <> 0 then %>
	     <td align="right" ><%=formatnumber(divamt2,2)%></td>
             <%else%>
             <td></td>
             <%end if%>
             <%if divamt3 <> 0 then %>
	     <td align="right" ><%=formatnumber(divamt3,2)%></td>
             <%else%>
             <td></td>
             <%end if%>
             <%if ajdivamt <> 0 then %>
             <td align="right" ><%=formatNUMBER(ajdivamt,2)%></td>
             <%else%>
             <td></td>
             <%end if%>
             <%if ttldiv <> 0 then %> 
	     <td align="right" ><%=formatnumber(ttldiv,2)%></td>
             <%else%>
             <td></td>
             <%end if%>
             <td></td>
	     <td></td>
        </tr>   
         <tr bgcolor="#FFFFFF">
             <td>  貸款　</td>
             <td align="right" ><%=formatnumber(actln,2)%></td>
	     <td></td>
	     <td></td>
             <td></td>	     
	     <td align="right" ><%=formatnumber(cloanamt,2)%></td>
             <td align="right" ><%=formatNUMBER(glnamt,2)%></td>
	     <td align="right" ><%=formatNUMBER(actlnamt,2)%></td>
        </tr>  
         <tr bgcolor="#FFFFFF">
             <td>  還款　</td>
             <td align="right" ><%=formatnumber(lbnkamt,2)%></td>
	     <td align="right" ><%=formatnumber(lsadamt,2)%></td>
	     <td align="right" ><%=formatnumber(lchamt,2)%></td>
             <td align="right" ><%=formatnumber(ajlnamt,2)%></td>	     
	     <td align="right" ><%=formatnumber(cpayamt,2)%></td>
             <td></td>
	     <td></td>
        </tr>  
        <tr bgcolor="#FFFFFF">
             <td>  利息　</td>
             <td align="right" ><%=formatnumber(ibnkamt,2)%></td>
	     <td align="right" ><%=formatnumber(isadamt,2)%></td>
	     <td align="right" ><%=formatnumber(ichamt,2)%></td>
             <td align="right" ><%=formatnumber(ajintamt,2)%></td>	     
	     <td align="right" ><%=formatnumber(cintamt,2)%></td>
             <td></td>
	     <td></td>
        </tr> 
       <tr bgcolor="#FFFFFF">
             <td>  入會費</td>
             <%if gamt1 <> 0 then %>
             <td align="right" ><%=formatnumber(gamt1,2)%></td>
             <%else%>
             <td></td>
             <%end if%>
             <%if gamt2 <> 0 then %>
	     <td align="right" ><%=formatnumber(gamt2,2)%></td>
             <%else%>
             <td></td>
             <%end if%>
             <%if gamt3 <> 0 then %>
	     <td align="right" ><%=formatnumber(gamt3,2)%></td>
             <%else%>
             <td></td>
             <%end if%>
             <td></td>	
             <%if ttlgamt <> 0 then %>     
	     <td align="right" ><%=formatnumber(ttlgamt,2)%></td>
             <%else%>
             <td></td>
             <%end if%>
             <td></td>
	     <td></td>
        </tr> 
      <tr bgcolor="#FFFFFF">
             <td>  協會費</td>
             <%if hamt1 <> 0 then %>
             <td align="right" ><%=formatnumber(hamt1,2)%></td>
             <%else%>
             <td></td>
             <%end if%>
             <%if hamt2 <> 0 then %>
	     <td align="right" ><%=formatnumber(hamt2,2)%></td>
             <%else%>
             <td></td>
             <%end if%>
             <%if hamt3 <> 0 then %>
	     <td align="right" ><%=formatnumber(hamt3,2)%></td>
             <%else%>
             <td></td>
             <%end if%>
             <td></td>	     
             <%if ttlHamtt <> 0 then %>
	     <td align="right" ><%=formatnumber(ttlHamtt,2)%></td>
             <%else%>
             <td></td>
             <%end if%>
             <td></td>
	     <td align="center">貸款/股金</td>
        </tr> 
      <tr bgcolor="#FFFFFF">
             <td>  收　入</td>
             <td align="right" ><%=formatnumber(ttlbnk,2)%></td>
	     <td align="right" ><%=formatnumber(ttlsad,2)%></td>
	     <td align="right" ><%=formatnumber(ttlch,2)%></td>
             <td align="right" ><%=formatnumber(cajttlamt,2)%></td>	     
	     <td align="right" ><%=formatnumber(ttlrec,2)%></td>
 
             <td align="center">總額</td>
             <td align="cener" ><%=formatnumber(gttlrate,2)%>%</td>
        </tr> 
        <BR>
     <tr bgcolor="#FFFFFF">
             <td>  支　出</td>
             <td align="right" ><%=formatnumber(payamt,2)%></td>
	     <td></td>
	     <td></td>
             <td align="right" ><%=formatnumber(ajwdamt,2)%></td>	     
	     <td align="right" ><%=formatnumber(ttlpay,2)%></td>
             <td align="right" ></td>
	     <td align="right" ></td>
        </tr> 
        <BR>
        <BR>

</table>
<table border="" cellpadding="0" cellspacing="0">
<tr>
<td>    貸款總數於<%=mm%>月<%=yy%>年前  </td> 
<td align="right"><%=formatnumber(ttllncnt,0)%></td>
</tr>
<tr>
<td>    新貸款總數</td> 
<td align="right">+<%=formatnumber(nwlncnt,0)%></td>
</tr>
<tr>
<td>    已清數貸款總數 </td>
<td align="right">-<%=formatnumber(clncnt,0)%></td> 
</tr>
<tr>
<td>    貸款總數合共   </td> 
<td align="right"><%=formatnumber(actlncnt,0)%></td>
</tr>
<tr></tr>
<tr></tr>
<tr>
<td>   社員總數於<%=mm%>月<%=yy%>年前 </td> 
<td align="right"><%=formatnumber(ttlmem,0)%></td>
</tr>
<tr>
<td>    社員退社總數 </td> 
<td align="right">-<%=formatnumber(omemcnt,0)%></td>
</tr>
<tr>
<td>    新社員總數</td> 
<td align="right">+<%=formatnumber(nmemcnt,0)%></td>
</tr>
<tr>
<td>    社員總數合共</td> 
<td align="right"><%=formatnumber(actmem,0)%></td>
</tr>

</table>
</center>
</body>
</html>

