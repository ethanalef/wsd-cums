<!-- #include file="../conn.asp" -->

<%
sqlstr=" and "
mndate=right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
if request.form("TS1")<>"" then
   sqlstr = sqlstr&"(  a.mstatus ='A'  "
end if

if request.form("TS2")<>  "" then
   if sqlstr=" and " then
      sqlstr= sqlstr&"( a.mstatus='0'  "
   else
      sqlstr= sqlstr & " or a.mstatus='0'  "
   end if
end if
if request.form("TS3")<>"" then
  if sqlstr=" and " then
      sqlstr=sqlstr& " ( a.mstatus='1'  "
   else
      sqlstr= sqlstr & "or a.mstatus='1' "
end if
end if
if request.form("TS4")<> "" then
   if sqlstr=" and " then
      sqlstr=sqlstr& " ( a.mstatus='2'  "
   else
      sqlstr= sqlstr & "or a.mstatus='2' "
   end if
end if
if request.form("TS5")<> "" then
   if sqlstr=" and " then
      sqlstr= sqlstr&" ( a.mstatus='3'  "
   else
      sqlstr= sqlstr & "or a.mstatus='3' "
   end if
end if
if request.form("TS6")<> "" then
   if sqlstr=" and " then
      sqlstr= sqlstr&"( a.mstatus='M'  "
   else
      sqlstr= sqlstr & "or a.mstatus='M' "
   end if
end if
if request.form("TS7")<> "" then
   if sqlstr=" and " then
      sqlstr=sqlstr& " ( a.mstatus='L'   "
   else
      sqlstr= sqlstr & "or a.mstatus='L' "
   end if
end if
if request.form("TS8")<> "" then
   if sqlstr=" and " then
      sqlstr=sqlstr& " ( a.mstatus='D'   "
   else
      sqlstr= sqlstr & "or a.mstatus='D' "
   end if
end if
if request.form("TS9")<> "" then
   if sqlstr=" and " then
      sqlstr=sqlstr& " ( a.mstatus='V'   "
   else
      sqlstr= sqlstr & "or a.mstatus='V' "
   end if
end if
if request.form("TS10")<> "" then
   if sqlstr=" and " then
      sqlstr= sqlstr&" ( a.mstatus='C'   "
   else
      sqlstr= sqlstr & "or a.mstatus='C' "
   end if
end if
if request.form("TS11")<> "" then
   if sqlstr=" and " then
      sqlstr= sqlstr&" ( a.mstatus='T'  "
   else
      sqlstr= sqlstr & "or a.mstatus='T' "
   end if
end if
if request.form("TS12")<> "" then
   if sqlstr=" and " then
      sqlstr= sqlstr&" ( a.mstatus='P'   "
   else
      sqlstr= sqlstr & "or a.mstatus='P' "
   end if
end if
if request.form("TS13")<> "" then
   if sqlstr=" and " then
      sqlstr= sqlstr&"( a.mstatus='B'  "
   else
      sqlstr= sqlstr & "or a.mstatus='B' "
   end if
end if
if request.form("TS14")<> "" then
   if sqlstr=" and " then
      sqlstr= sqlstr&" ( a.mstatus='N'   "
   else
      sqlstr= sqlstr & "or a.mstatus='N' "
   end if
end if

if request.form("TS15")<> "" then
   if sqlstr=" and " then
      sqlstr= sqlstr&" ( a.mstatus='J'   "
   else
      sqlstr= sqlstr & "or a.mstatus='J' "
   end if
end if
if request.form("TS16")<> "" then
   if sqlstr=" and " then
      sqlstr= sqlstr&" ( a.mstatus='H'   "
   else
      sqlstr= sqlstr & "or a.mstatus='H' "
   end if
end if
if request.form("TS17")<> "" then
   if sqlstr=" and " then
      sqlstr= sqlstr&" ( a.mstatus='F'   "
   else
      sqlstr= sqlstr & "or a.mstatus='F' "
   end if
end if
if request.form("TS18")<> "" then
   if sqlstr=" and " then
      sqlstr= sqlstr&" ( a.mstatus='8'  "
   else
      sqlstr= sqlstr & "or a.mstatus='8' "
   end if
end if
if request.form("TS19")<> "" then
   if sqlstr=" and " then
      sqlstr= sqlstr&" ( a.mstatus='9'  "
   else
      sqlstr= sqlstr & "or a.mstatus='9' "
   end if
end if

if request.form("TS20")<> "" then
       sqlstr ="   GROUP BY a.memno,a.memcname,a.membday,a.memhkid,a.memgender,a.mstatus,B.CODE order by a.memno,a.memcname,a.membday,a.memhkid,a.memgender,a.mstatus,B.CODE"
else
  if sqlstr=" and " then
     response.redirect "memstlst.asp"
  else
     sqlstr = sqlstr&" )  GROUP BY a.memno,a.memcname,a.membday ,a.memhkid,a.memgender,a.mstatus,B.CODE order by a.memno,a.memcname,a.membday,a.memhkid,a.memgender,a.mstatus,B.CODE"
  end if
end if

sxdate=request.form("cutdate")

if request.form("TS8")<> "" then
    mdate= dateserial(right(sxdate,4),7,1 )
else
  mdate= dateserial(right(sxdate,4),mid(sxdate,4,2),left(sxdate,2))
end if

set rs = server.createobject("ADODB.Recordset")

sql = "select a.memno,a.memcname,a.membday,a.memhkid,a.memgender,a.mstatus,B.CODE,SUM(b.amount) "&_
      "from memmaster a,share b  where   a.memno=b.memno and  b.ldate<'"&mdate&"' "&sqlstr

rs.open sql, conn, 1, 1

if rs.eof then
   response.redirect "noprint.asp"
end if

ttlamt = 0

if request.form("output")="word" then
  Response.ContentType = "application/msword"
elseif request.form("output")="excel" then
  Response.ContentType = "application/vnd.ms-excel"

end if
%>
<html>
<head>
<title>社員狀況列印</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
<center>

<table width="1012" border="0">

 <tr>
    <td width="99">&nbsp;</td>
    <td width="780">&nbsp;</td>
    <td width="142">&nbsp;</td>
  </tr>
  <tr>
        <td>&nbsp</td>
        <td align="center"><b><font size="4"  face="標楷體" >水務署員工儲蓄互助社<br>社員狀況列印</font></b?</td>
        <td align="center"><font size="2"  face="標楷體" >日期 : <%=mndate%></font></td>
        </tr>


</table>
<table border="0" cellpadding="0" cellspacing="0">
  <tr height="15" valign="bottom">
  <td width="80" align="center"><font size="3" face="標楷體"  >社員編號</font></td>
  <td width="80"  align="center"><font size="3" face="標楷體"  > 姓名</font></td>
  <td width="80"  align="center"><font size="3" face="標楷體"  > 身分證</font></td>
  <td width="80"  align="center"><font size="3" face="標楷體"  > 出生日期</font></td>
  <td width="80"  align="center"><font size="3" face="標楷體"  > 性別</font></td>
  <td width="80"  align="center"><font size="3" face="標楷體"  > 年齡</font></td>
  <td width="130" align="center"><font size="3" face="標楷體"  >股金結餘</font></td>
        <td width="130" align="center"><font size="3" face="標楷體"  >貸款結餘</font></td>
        <td width="130" align="center"><font size="3" face="標楷體"  >分類</font></td>
  </tr>
  <tr><td colspan=10><hr></td></tr>
<%
   tlnamt = 0

   ttlamt = 0
   ttl1   = 0
   ttl2   =0
   ttl3   =0
   ttl4   =0
   ttl5   = 0
   ttl6   =0
   ttl7   =0
   ttl8   =0
   ttl9   = 0
   ttl10   =0
   ttl11   =0
   ttl12   =0
   ttl13   = 0
   ttl14   =0
   ttl15   =0
   ttl16   =0
   ttl17   = 0
   ttl18   =0
   ttl19   =0
   cnt1   = 0
   cnt2   =0
   cnt3   =0
   cnt4   =0
   cnt5   = 0
   cnt6   =0
   cnt7   =0
   cnt8   =0
   cnt9   = 0
   cnt10   =0
   cnt11   =0
   cnt12   =0
   cnt13   = 0
   cnt14   =0
   cnt15   =0
   cnt16   =0
   cnt17   = 0
   cnt18   =0
   cnt19   =0
   xamt = 0
   lnamt = 0
   clsbal = 0
   ttlamt = 0
   ttllncnt = 0
   xmemno=rs("memno")
   memcname=rs("memcname")
   xmstatus=  rs("mstatus")
   xmembday = rs("membday")
   xmemhkid = rs("memhkid")
   xmemgender =rs("memgender")
   if xmembday<>"" then
         membday = right("0"&day(xmembday),2)&"/"&right("0"&month(xmembday),2)&"/"&year(xmembday)
                  age =year(date()) -year(xmembday)
                  yy = year(date())
                  adatae = dateserial(yy,1,1)
                  mm = month(xmembday)
                  dd = day(xmembday)

                  bdate = dateserial(yy,mm,dd)
                  xday = ((bdate - adate)+z)/365
                  if xday >0.5 then
                     age = age + 1
                  end if
 else
       age = 0
       membday=""
  end if
      if xmemGender="M" then

         sex="先生"
      else

          sex="女士"
      end if
   do while not rs.eof



      if xmemno<> rs("memno") then

         xamt = 0
         lnamt = 0
         ylnnum=""
         set ms = server.createobject("ADODB.Recordset")
         mssql = "select lnnum,appamt,bal from loanrec where memno ='"&xmemno&"' and repaystat='N'   "
         ms.open mssql, conn, 1, 1
         if not ms.eof then

            lnamt  = ms("bal")

         end if

         ms.close

         if lnamt > 0 then
            tlnamt = tlnamt + lnamt
            ttllncnt = ttllncnt + 1

         end if
         select case xmstatus
                case "A"
                     if clsbal > 0 then
                        ttl1 = ttl1 + clsbal
                        cnt1 = cnt1 + 1
                     end if
                     idx ="自動轉帳(ALL)"
                case "B"
                     if clsbal > 0  or (clsbal=0 and lnamt >0)  then
                        ttl2 = ttl2 +clsbal
                        cnt2 = cnt2 + 1
                      end if

                     idx ="破產"
                case "C"
                   if clsbal > 0  or (clsbal=0 and lnamt >0)  then
                       ttl3 = ttl3+ clsbal
                       cnt3 = cnt3 + 1
                     end if
                     idx ="退社"
                case "D"
                    if clsbal > 0 then
                       ttl4 = ttl4 + clsbal
                       cnt4 = cnt4 + 1
                    end if
                     idx ="冷戶"
                case "F"
                    ttl5 = ttl5 +  clsbal
                    cnt5 = cnt5 + 1
                     idx ="特別個案"
                case "H"
                   ttl6 = ttl6 + clsbal
                   cnt6 = cnt6 + 1
                     idx ="暫停銀行"
                case  "J"
                   ttl7 = ttl7 + clsbal
                   cnt7 = cnt7 + 1
                     idx ="新戶"
                case "L"
                 if clsbal > 0 or (clsbal=0 and lnamt>0) then
                   ttl8 = ttl8 + clsbal
                   cnt8 = cnt8 + 1
                 end if
                     idx ="呆帳"
                case "M"
                  ttl9 = ttl9 + clsbal
                  cnt9 = cnt9 + 1
                     idx ="庫房,銀行"
                case "N"
                 ttl10 = ttl10 + clsbal
                 cnt10 = cnt10 + 1
                     idx ="正常"
                case "P"
                if clsbal > 0 then
                   ttl11 = ttl11 + clsbal
                   cnt11 = cnt11 + 1
                end if
                     idx ="去世"
                case "T"
                 ttl12 = ttl12 + clsbal
                 cnt12 = cnt12 + 1
                     idx ="庫房"
                case "V"
                     if clsbal> 0 or (clsbal=0 and lnamt >0 ) then
                        ttl13 = ttl13 + clsbal
                        cnt13 = cnt13 + 1
                     end if
                     idx ="IVA"

                case "0"
                 ttl14 = ttl14 + clsbal
                 cnt14 = cnt14 + 1
                     idx =" 自動轉帳(股金)"
                case "1"
                 ttl15 = ttl15 + clsbal
                 cnt15 = cnt15 + 1
                     idx ="自動轉帳(股金,利息)"
                case "2"
                 ttl16 = ttl16 + clsbal
                 cnt16 = cnt16 + 1
                     idx ="自動轉帳(股金,本金)"
                case "3"
                 ttl17 = ttl17 + clsbal
                 cnt17 = cnt17 + 1
                     idx ="自動轉帳(利息,本金)"
                case "8"
                 ttl18 = ttl18 + clsbal
                 cnt18 = cnt18 + 1
                     idx ="終止社籍轉帳"
                case "9"
                 ttl19 = ttl19 + clsbal
                 cnt19 = cnt19 + 1
                     idx ="終止社籍正常"
         end select
    if clsbal > 0 or (clsbal=0 and lnamt > 0 ) or xmstatus = "C" then
%>
     <tr>
          <td width="80"><%=xmemno%></td>
          <td width="80"><font size="3" face="標楷體"  ><%=memcname%></font></td>
          <td width="80"><%=xmemhkid%></td>
          <td width="80"><%=membday%></td>
          <td width="80" align="center"><%=sex%></td>
          <td width="80"  align="center"><%=formatnumber(age,0)%></td>
          <td width="130" align="right"><%=formatnumber(clsbal,2)%></td>
          <td width="130" align="right"><%=formatnumber(lnamt,2)%></td>
          <td width="150" align="center"><font size="3" face="標楷體"  ></font><%=idx%></font></td>
     </tr>

<%
      ttlamt = ttlamt + clsbal
   end if
        xmemno=rs("memno")
          memcname=rs("memcname")
          xmstatus=rs("mstatus")
  xmembday = rs("membday")
   xmemhkid = rs("memhkid")
   xmemgender =rs("memgender")
  if xmembday<>""   then
        membday = right("0"&day(xmembday),2)&"/"&right("0"&month(xmembday),2)&"/"&year(xmembday)
                  age =year(date()) -year(xmembday)
                  yy = year(date())
                  adatae = dateserial(yy,1,1)
                  mm = month(xmembday)
                  dd = day(xmembday)

                  bdate = dateserial(yy,mm,dd)
                  xday = ((bdate - adate)+z)/365
                  if xday >0.5 then
                     age = age + 1
                  end if
  else
       age = 0
       membday=""
  end if
      if xmemGender="M" then

         sex="先生"
      else

          sex="女士"
      end if
        clsbal = 0
        lnamt = 0
        xamt = 0
   end if


                            select case rs("code")
                                    case "A1","A2","A3","C0","C1","C3" ,"A0","A7" ,"A4" ,"C5","0A"
                                         clsbal = clsbal + rs(7)
                                    case "G0" ,"H0","B0","B1","B3","BE","BF","G3","H3","MF"
                                         clsbal= clsbal - rs(7)
                             end select



     rs.movenext
    loop

         xamt = 0
         lnamt = 0
         ylnnum=""
         set ms = server.createobject("ADODB.Recordset")
         mssql = "select lnnum,appamt,bal from loanrec where memno ='"&xmemno&"' and repaystat='N'   "
         ms.open mssql, conn, 1, 1
         if not ms.eof then

            lnamt  = ms("bal")

         end if

         ms.close

         if lnamt > 0 then
            tlnamt = tlnamt + lnamt
            ttllncnt = ttllncnt + 1
         end if
          select case xmstatus
                case "A"
                     if clsbal > 0 then
                        ttl1 = ttl1 + clsbal
                        cnt1 = cnt1 + 1
                     end if
                     idx ="自動轉帳(ALL)"
                case "B"
                     if clsbal > 0  or (clsbal=0 and lnamt >0)  then
                        ttl2 = ttl2 +clsbal
                        cnt2 = cnt2 + 1
                      end if

                     idx ="破產"
                case "C"
                   if clsbal > 0  or (clsbal=0 and lnamt >0)  then
                       ttl3 = ttl3+ clsbal
                       cnt3 = cnt3 + 1
                     end if
                     idx ="退社"
                case "D"
                    if clsbal > 0 then
                       ttl4 = ttl4 + clsbal
                       cnt4 = cnt4 + 1
                    end if
                     idx ="冷戶"
                case "F"
                    ttl5 = ttl5  + clsbal
                    cnt5 = cnt5 + 1
                     idx ="特別個案"
                case "H"
                   ttl6 = ttl6 + clsbal
                   cnt6 = cnt6 + 1
                     idx ="暫停銀行"
                case  "J"
                   ttl7 = ttl7 + clsbal
                   cnt7 = cnt7 + 1
                     idx ="新戶"
                case "L"
                  if clsbal > 0 or (clsbal=0 and lnamt>0) then
                   ttl8 = ttl8 + clsbal
                   cnt8 = cnt8 + 1
                 end if
                     idx ="呆帳"
                case "M"
                  ttl9 = ttl9 + clsbal
                  cnt9 = cnt9 + 1
                     idx ="庫房,銀行"
                case "N"
                 ttl10 = ttl10 + clsbal
                 cnt10 = cnt10 + 1
                     idx ="正常"
                case "P"
                if clsbal > 0 then
                   ttl11 = ttl11 + clsbal
                   cnt11 = cnt11 + 1
                end if
                     idx ="去世"
                case "T"
                 ttl12 = ttl12 + clsbal
                 cnt12 = cnt12 + 1
                     idx ="庫房"
                case "V"
                      if clsbal> 0 or (clsbal=0 and lnamt >0 ) then
                        ttl13 = ttl13 + clsbal
                        cnt13 = cnt13 + 1
                     end if
                     idx ="IVA"
                case "0"
                 ttl14 = ttl14 + clsbal
                 cnt14 = cnt14 + 1
                     idx =" 自動轉帳(股金)"
                case "1"
                 ttl15 = ttl15 + clsbal
                 cnt15 = cnt15 + 1
                     idx ="自動轉帳(股金,利息)"
                case "2"
                 ttl16 = ttl16 + clsbal
                 cnt16 = cnt16 + 1
                     idx ="自動轉帳(股金,本金)"
                case "3"
                 ttl17 = ttl17 + clsbal
                 cnt17 = cnt17 + 1
                     idx ="自動轉帳(利息,本金)"
                case "8"
                 ttl18 = ttl18 + clsbal
                 cnt18 = cnt18 + 1
                     idx ="終止社籍轉帳"
                case "9"
                 ttl19 = ttl19 + clsbal
                 cnt19 = cnt19 + 1
                     idx ="終止社籍正常"
         end select
   if clsbal > 0 or (clsbal=0 and lnamt > 0 ) then
 %>
     <tr>
          <td width="80"><%=xmemno%></td>
          <td width="80"><font size="3" face="標楷體"  ><%=memcname%></font></td>
        <td width="80"><%=xmemhkid%></td>
          <td width="80"><%=membday%></td>
          <td width="80" align="center"><%=sex%></td>
          <td width="80"  align="center"><%=formatnumber(age,0)%></td>
          <td width="130" align="right"><%=formatnumber(clsbal,2)%></td>
          <td width="130" align="right"><%=formatnumber(lnamt,2)%></td>
          <td width="150" align="center"><font size="3" face="標楷體"  ><%=idx%></font></td>
     </tr>

<%
        ttlamt = ttlamt + clsbal

 end if
 %>

       <tr><td colspan=10><hr></td></tr>
        <tr>
              <td></td>
              <td></td>
              <td></td>
              <td></td>
              <td></td>
             <td></td>
             <td width="130" align="right"><%=formatnumber(ttlamt,2)%></td>
             <td width="130" align="right"><%=formatnumber(tlnamt,2)%></td>
         </tr>
        <tr><td></td>
             <td></td>
              <td></td>
              <td></td>
              <td></td>
             <td></td>
             <td width="130" align="right">============</td>
              <td width="130" align="right">============</td>
         </tr>


</table>
<BR>

<BR>

<table border="0" cellpadding="0" cellspacing="0">
<tr>
     <td width="200" ><font size="3" face="標楷體"  >貸款合約合共 :</font></td>
      <td width="100" align="right"><%=formatNumber(ttllncnt,0)%></td>
      <td width="30">
       <td></td>
       <td></td>
</tr>
<%
     ttlcnt =         cnt1+ cnt2+ cnt3+ cnt4+ cnt5+ cnt6+ cnt7+ cnt8+ cnt9+cnt10
     ttlcnt = ttlcnt +cnt11+cnt12+cnt13+cnt14+cnt15+cnt16+cnt17+cnt18+cnt19

     if ttl1 <> 0 then %>
<tr>
      <td width="200" ><font size="3" face="標楷體"  >自動轉帳(ALL)金額合共 :</font></td>
      <td width="100" align="right"><%=formatNumber(ttl1,2)%></td>
      <td width="30">
      <td width="200" ><font size="3" face="標楷體"  >自動轉帳(ALL人數合共 :</font></td>
      <td width="100" align="right"><%=formatNumber(cnt1,0)%></td>
</tr>
<%end if%>
<% if ttl2 <> 0 then %>
 <tr>
      <td width="200" ><font size="3" face="標楷體"  >破產金額合共 :</font></td>
      <td width="100" align="right"><%=formatNumber(ttl2,2)%></td>
      <td width="30">
      <td width="200" >破產人數合共 :</td>
      <td width="100" align="right"><%=formatNumber(cnt2,0)%></font></td>
</tr>
<%end if%>
<% if ttl3 <> 0 then %>
 <tr>
      <td width="200" >退社金額合共 :</td>
      <td width="100" align="right"><%=formatNumber(ttl3,2)%></font></td>
      <td width="30">
      <td width="200" ><font size="3" face="標楷體"  >退社人數合共 :</font></td>
      <td width="100" align="right"><%=formatNumber(cnt3,0)%></td>
</tr>
<%end if%>
<% if ttl4 <> 0 then %>
<tr>
      <td width="200" ><font size="3" face="標楷體"  >冷戶金額合共 :</font></td>
      <td width="100" align="right"><%=formatNumber(ttl4,2)%></td>
      <td width="30">
      <td width="200" ><font size="3" face="標楷體"  >冷戶人數合共 :</font></td>
      <td width="100" align="right"><%=formatNumber(cnt4,0)%></td>
</tr>
<%end if%>
<% if ttl5 <> 0 then %>
 <tr>
      <td width="200" ><font size="3" face="標楷體"  >特別個案金額合共 :</font></td>
      <td width="100" align="right"><%=formatNumber(ttl5,2)%></td>
      <td width="30">
      <td width="200" ><font size="3" face="標楷體"  >特別個案人數合共 :</font></td>
      <td width="100" align="right"><%=formatNumber(cnt5,0)%></td>
</tr>
<%end if%>
<% if ttl6 <> 0 then %>
 <tr>
      <td width="200" ><font size="3" face="標楷體"  >暫停銀行金額合共 :</font></td>
      <td width="100" align="right"><%=formatNumber(ttl6,2)%></td>
      <td width="30">
      <td width="200" ><font size="3" face="標楷體"  >暫停銀行人數合共 :</font></td>
      <td width="100" align="right"><%=formatNumber(cnt6,0)%></td>
</tr>
<%end if%>
<% if ttl7 <> 0 then %>
<tr>
      <td width="200" ><font size="3" face="標楷體"  >新戶金額合共 :</font></td>
      <td width="100" align="right"><%=formatNumber(ttl7,2)%></td>
      <td width="30">
      <td width="200" ><font size="3" face="標楷體"  >新戶人數合共 :</font></td>
      <td width="100" align="right"><%=formatNumber(cnt7,0)%></td>
</tr>
<%end if%>
<% if ttl8 <> 0 then %>
 <tr>
      <td width="200" ><font size="3" face="標楷體"  >呆帳金額合共 :</font></td>
      <td width="100" align="right"><%=formatNumber(ttl8,2)%></td>
      <td width="30">
      <td width="200" ><font size="3" face="標楷體"  >呆帳人數合共 :</font></td>
      <td width="100" align="right"><%=formatNumber(cnt8,0)%></td>
</tr>
<%end if%>
<% if ttl9 <> 0 then %>
 <tr>
      <td width="200" ><font size="3" face="標楷體"  >庫房,銀行金額合共 :</font></td>
      <td width="100" align="right"><%=formatNumber(ttl9,2)%></td>
      <td width="30">
      <td width="200" ><font size="3" face="標楷體"  >庫房,銀行人數合共 :</font></td>
      <td width="100" align="right"><%=formatNumber(cnt9,0)%></td>
</tr>
<%end if%>
<% if ttl10 <> 0 then %>
<tr>
      <td width="200" ><font size="3" face="標楷體"  >正常金額合共 :</font></td>
      <td width="100" align="right"><%=formatNumber(ttl10,2)%></td>
      <td width="30">
      <td width="200" ><font size="3" face="標楷體"  >正常人數合共 :</font></td>
      <td width="100" align="right"><%=formatNumber(cnt10,0)%></td>
</tr>
<%end if%>
<% if ttl11 <> 0 then %>
 <tr>
      <td width="200" ><font size="3" face="標楷體"  >去世金額合共 :</font></td>
      <td width="100" align="right"><%=formatNumber(ttl11,2)%></td>
      <td width="30">
      <td width="200" ><font size="3" face="標楷體"  >去世人數合共 :</font></td>
      <td width="100" align="right"><%=formatNumber(cnt11,0)%></td>
</tr>
<%end if%>
<% if ttl12 <> 0 then %>
 <tr>
      <td width="200" ><font size="3" face="標楷體"  >庫房金額合共 :</font></td>
      <td width="100" align="right"><%=formatNumber(ttl12,2)%></td>
      <td width="30">
      <td width="200" ><font size="3" face="標楷體"  >庫房人數合共 :</font></td>
      <td width="100" align="right"><%=formatNumber(cnt12,0)%></td>
</tr>
<%end if%>

<% if ttl13 <> 0 then %>
 <tr>
      <td width="200" ><font size="3" face="標楷體"  >IVA金額合共 :</font></td>
      <td width="100" align="right"><%=formatNumber(ttl13,2)%></td>
      <td width="30">
      <td width="200" ><font size="3" face="標楷體"  >IVA人數合共 :</font></td>
      <td width="100" align="right"><%=formatNumber(cnt13,0)%></td>
</tr>
<%end if%>

<% if ttl14 <> 0 then %>
<tr>
      <td width="200" ><font size="3" face="標楷體"  > 自動轉帳(股金)金額合共 :</font></td>
      <td width="100" align="right"><%=formatNumber(ttl14,2)%></td>
      <td width="30">
      <td width="200" > <font size="3" face="標楷體"  >自動轉帳(股金)人數合共 :</font></td>
      <td width="100" align="right"><%=formatNumber(cnt14,0)%></td>
</tr>
<%end if%>
<% if ttl15 <> 0 then %>
 <tr>
      <td width="200" ><font size="3" face="標楷體"  >自動轉帳(股金,利息)金額合共 :</font></td>
      <td width="100" align="right"><%=formatNumber(ttl15,2)%></td>
      <td width="30">
      <td width="200" ><font size="3" face="標楷體"  >自動轉帳(股金,利息)人數合共 :</font></td>
      <td width="100" align="right"><%=formatNumber(cnt15,0)%></td>
</tr>
<%end if%>
<% if ttl16 <> 0 then %>
 <tr>
      <td width="200" ><font size="3" face="標楷體"  >自動轉帳(股金,本金)金額合共 :</font></td>
      <td width="100" align="right"><%=formatNumber(ttl16,2)%></td>
      <td width="30">
      <td width="200" ><font size="3" face="標楷體"  >自動轉帳(股金,本金)人數合共 :</font></td>
      <td width="100" align="right"><%=formatNumber(cnt16,0)%></td>
</tr>
<%end if%>
<% if ttl17 <> 0 then %>
 <tr>
      <td width="200" ><font size="3" face="標楷體"  >自動轉帳(利息,本金)金額合共 :</font></td>
      <td width="100" align="right"><%=formatNumber(ttl17,2)%></td>
      <td width="30">
      <td width="200" ><font size="3" face="標楷體"  >自動轉帳(利息,本金)人數合共 :</font></td>
      <td width="100" align="right"><%=formatNumber(cnt17,0)%></td>
</tr>
<%end if%>
<% if ttl18 <> 0 then %>
 <tr>
      <td width="200" ><font size="3" face="標楷體"  >終止社籍轉帳金額合共 :</font></td>
      <td width="100" align="right"><%=formatNumber(ttl18,2)%></td>
      <td width="30">
      <td width="200" ><font size="3" face="標楷體"  >終止社籍轉帳人數合共 :</font></td>
      <td width="100" align="right"><%=formatNumber(cnt18,0)%></td>
</tr>
<%end if%>
<% if ttl19 <> 0 then %>
 <tr>
      <td width="200" ><font size="3" face="標楷體"  >終止社籍正常金額合共 :</font></td>
      <td width="100" align="right"><%=formatNumber(ttl19,2)%></td>
      <td width="30">
      <td width="200" ><font size="3" face="標楷體"  >終止社籍正常人數合共 :</font></td>
      <td width="100" align="right"><%=formatNumber(cnt19,0)%></td>
</tr>
<%end if%>
  <tr><td colspan=5><hr></td></tr>
        <tr><td></td>
            <td width=100 align="right"><%=formatnumber(ttlamt,2)%></td>
      <td></td>
            <td></td>
            <td width=100 align="right"><%=formatnumber(ttlcnt,0)%></td>
            <td></td>
        </tr>
        <tr>
            <td width=200 align="right"></td>
            <td width=100 align="right">===========</td>
               <td width="30">
            <td width=200 align="right"></td>
             <td width=100 align="right">======</td>


        </tr>
</table>
</center>
</body>
</html>

