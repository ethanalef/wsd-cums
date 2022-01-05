<!-- #include file="../conn.asp" -->

<%

server.scripttimeout = 1800

SQl = "SELECT  a.memno,a.adate,sum(a.bankin) as unpaid ,b.memname,b.memcname,b.accode  FROM  autopay a ,memmaster b where a.memno=b.memno and a.flag='F' and right(a.code,1)='1' and a.pflag=1 group by a.memno,a.adate,b.memname,b.memcname,b.accode  order by a.memno,a.adate,b.memname,b.memcname,b.accode  "
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn
if rs.eof then
   response.redirect "rejectlst.asp"
end if
dim guarantor(3)
dim gender(3)
if request.form("output")="word" then
	Response.ContentType = "application/msword"
        elseif request.form("output")="excel" then
	Response.ContentType = "application/vnd.ms-excel"
end if
%>
<html>
<head>
<title>銀行轉賬失效通知書</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body leftMargin="0" topmargin="0" marginheight="0" marginwidth="0">
<table width="1012" border="0">

 <tr>
    <td width="99">&nbsp;</td>
    <td width="780">&nbsp;</td>
    <td width="142">&nbsp;</td>
  </tr>
<%
ttlcnt = 0
do while not rs.eof
for i = 1 to 3
    guarantor(i)=""
    gender(i) =""
next
xx = 1
sqlstr = "select a.*,b.memGender from   guarantor a,memmaster b where a.memno='"&rs("memno")&"' and a.memno=b.memno "
Set ms = Server.CreateObject("ADODB.Recordset")
ms.open sqlstr, conn,2,2
if not ms.eof then
   do while not ms.eof
 
      guarantor(xx)= ms("guarantorCname")   
      if ms("memGender")="M" then
         guarantor(xx) = guarantor(xx)
         gender(xx)="先生"
      else
          guarantor(xx) = guarantor(xx)
          gender(i)="女士"
      end if
      ms.movenext
   loop
end if
ms.close
   yy = right(year(rs("adate")),2)
   mm=month(rs("adate"))  
   ttlcnt = ttlcnt + 1   
   refno="AR"&yy&right("0"&mm,2)&right("0000"&ttlcnt,4)

%>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
 
 <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>

 
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><p><%=refno%></p></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><font size="3" face="標楷體" > 親愛的社員 ：　<%=rs("memcname")%> </font></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><font size="3" face="標楷體" > 社員編號　：　<%=rs("memno")%> </font></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>　　　　　　　　　　　　　<u><font size="3" face="標楷體" >按月銀行自動轉帳失效通知書</font></u></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><p><font size="3" face="標楷體" > 　　　　多謝你一直以來對本社的支持和信任！ </font></p></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><font size="3" face="標楷體" > 　　　　依據本社記錄顯示，從銀行覆函得知，本社於上月底未能從閣下戶口收取 </font></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><font size="3" face="標楷體" > 該月應繳納之款項，相信可能是一時忘記。請閣下於接獲此通知書後，從速聯絡本 </font></td>
    <td>&nbsp;</td>
  </tr>
  <tr>

    <td>&nbsp;</td>
<% if rs("accode")=9999 then %>
    <td><font size="3" face="標楷體" > 社辦事處安排補回款項 $<%=formatnumber(rs("unpaid"),2)%>。敝社定當盡力協助閣下解決有關財務事宜，以</font></td>

<%else%>
    <td><font size="3" face="標楷體" > 社辦事處或各分區委員安排補回款項 $<%=formatnumber(rs("unpaid"),2)%>。敝社定當盡力協助閣下解決有關 </font></td>
<%end if %>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
<% if rs("accode")=9999 then %>
    <td><p><font size="3" face="標楷體" >免影響日後借貸信用或其他利息上之損失。 </font></p></td>
<%else%>
    <td><p><font size="3" face="標楷體" >財務事宜，以免影響日後借貸信用或其他利息上之損失。 </font></p></td>
<%end if %>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><font size="3" face="標楷體" > 　　　　本社不鼓勵社員經常發生自動轉帳和還款脫期等情況；而為了加強本社之 </font></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><font size="3" face="標楷體" > 工作效率與及保障其他社員權益，我們會盡量在不滋擾閣下工作的情況下，經常及 </font></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><font size="3" face="標楷體" > 第一時間提醒閣下與及閣下之擔保人（如適用）有關上述事宜，不便之處，敬請原 </font></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><font size="3" face="標楷體" > 諒！</font> </td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><font size="3" face="標楷體" > 　　　　假如閣下已從其他方式，例如現金、支票或過戶形式，補回上述之款項， </font></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><p><font size="3" face="標楷體" >則無須理會此通知書。 </font></p></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><p><font size="3" face="標楷體" >　　　　如有任何查詢，歡迎致電 2787 9222 與本社職員聯絡。 </font></p></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><p><font size="3" face="標楷體" >水務署員工儲蓄互助社 </font></p></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><p><font size="3" face="標楷體" >董事會 司庫 </font></p></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><%=year(date())%><font size="3" face="標楷體" >年</font><%=month(date())%><font size="3" face="標楷體" >月</font><%=day(date())%><font size="3" face="標楷體" >日</font></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><p><font size="3" face="標楷體" >副本呈交 </font></p></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><p><font size="3" face="標楷體" >　　呆帳及冷戶管理小組</font></p></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><p><font size="3" face="標楷體" >　　貸款委員會 </font></p></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
<%if guarantor(1)<>"" then%>
    <td><p><font size="3" face="標楷體" >　　＊＊擔保人<u><%=guarantor(1)%></u><%=gender(1)%>＊＊</font></p></td>
<%else%>
    <td>&nbsp;</td>
<%end if%>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
<%if guarantor(2)<>"" then%>
    <td><p><font size="3" face="標楷體" >　　＊＊擔保人<u><%=guarantor(2)%></u><%=gender(2)%>＊＊</font></p></td>
<%else%>
    <td>&nbsp;</td>
<%end if%>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
<%if guarantor(3)<>"" then%>
    <td><p><font size="3" face="標楷體" >　　＊＊擔保人<u><%=guarantor(3)%></u><%=gender(3)%>＊＊</font></p></td>
<%else%>
    <td>&nbsp;</td>
<%end if%>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><p><em><font size="3" face="標楷體" >本函乃由電腦集中印發，無須簽署 </font></em><em></em></p></td>
    <td>&nbsp;</td>
  </tr>

  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>



<%
  RS.MOVENEXT
  LOOP
%> 
</font>
</table>
</body>
</html>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
