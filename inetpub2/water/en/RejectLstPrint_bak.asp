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
<title>�Ȧ���㥢�ĳq����</title>
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
         gender(xx)="����"
      else
          guarantor(xx) = guarantor(xx)
          gender(i)="�k�h"
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
    <td><font size="3" face="�з���" > �˷R������ �G�@<%=rs("memcname")%> </font></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><font size="3" face="�з���" > �����s���@�G�@<%=rs("memno")%> </font></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>�@�@�@�@�@�@�@�@�@�@�@�@�@<u><font size="3" face="�з���" >����Ȧ�۰���b���ĳq����</font></u></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><p><font size="3" face="�з���" > �@�@�@�@�h�§A�@���H�ӹ糧��������M�H���I </font></p></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><font size="3" face="�з���" > �@�@�@�@�̾ڥ����O����ܡA�q�Ȧ��Ш�o���A������W�멳����q�դU��f���� </font></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><font size="3" face="�з���" > �Ӥ���ú�Ǥ��ڶ��A�۫H�i��O�@�ɧѰO�C�лդU���򦹳q���ѫ�A�q�t�p���� </font></td>
    <td>&nbsp;</td>
  </tr>
  <tr>

    <td>&nbsp;</td>
<% if rs("accode")=9999 then %>
    <td><font size="3" face="�з���" > ����ƳB�w�Ƹɦ^�ڶ� $<%=formatnumber(rs("unpaid"),2)%>�C�ͪ��w��ɤO��U�դU�ѨM�����]�ȨƩy�A�H</font></td>

<%else%>
    <td><font size="3" face="�з���" > ����ƳB�ΦU���ϩe���w�Ƹɦ^�ڶ� $<%=formatnumber(rs("unpaid"),2)%>�C�ͪ��w��ɤO��U�դU�ѨM���� </font></td>
<%end if %>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
<% if rs("accode")=9999 then %>
    <td><p><font size="3" face="�з���" >�K�v�T���ɶU�H�ΩΨ�L�Q���W���l���C </font></p></td>
<%else%>
    <td><p><font size="3" face="�з���" >�]�ȨƩy�A�H�K�v�T���ɶU�H�ΩΨ�L�Q���W���l���C </font></p></td>
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
    <td><font size="3" face="�з���" > �@�@�@�@���������y�����g�`�o�ͦ۰���b�M�ٴڲ�������p�F�Ӭ��F�[�j������ </font></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><font size="3" face="�з���" > �u�@�Ĳv�P�ΫO�٨�L�����v�q�A�ڭ̷|�ɶq�b�����Z�դU�u�@�����p�U�A�g�`�� </font></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><font size="3" face="�з���" > �Ĥ@�ɶ������դU�P�λդU����O�H�]�p�A�Ρ^�����W�z�Ʃy�A���K���B�A�q�Э� </font></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><font size="3" face="�з���" > �̡I</font> </td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><font size="3" face="�з���" > �@�@�@�@���p�դU�w�q��L�覡�A�Ҧp�{���B�䲼�ιL��Φ��A�ɦ^�W�z���ڶ��A </font></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><p><font size="3" face="�з���" >�h�L���z�|���q���ѡC </font></p></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><p><font size="3" face="�з���" >�@�@�@�@�p������d�ߡA�w��P�q 2787 9222 �P����¾���p���C </font></p></td>
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
    <td><p><font size="3" face="�з���" >���ȸp���u�x�W���U�� </font></p></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><p><font size="3" face="�з���" >���Ʒ| �q�w </font></p></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><%=year(date())%><font size="3" face="�з���" >�~</font><%=month(date())%><font size="3" face="�з���" >��</font><%=day(date())%><font size="3" face="�з���" >��</font></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><p><font size="3" face="�з���" >�ƥ��e�� </font></p></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><p><font size="3" face="�з���" >�@�@�b�b�ΧN��޲z�p��</font></p></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><p><font size="3" face="�з���" >�@�@�U�کe���| </font></p></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
<%if guarantor(1)<>"" then%>
    <td><p><font size="3" face="�з���" >�@�@������O�H<u><%=guarantor(1)%></u><%=gender(1)%>����</font></p></td>
<%else%>
    <td>&nbsp;</td>
<%end if%>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
<%if guarantor(2)<>"" then%>
    <td><p><font size="3" face="�з���" >�@�@������O�H<u><%=guarantor(2)%></u><%=gender(2)%>����</font></p></td>
<%else%>
    <td>&nbsp;</td>
<%end if%>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
<%if guarantor(3)<>"" then%>
    <td><p><font size="3" face="�з���" >�@�@������O�H<u><%=guarantor(3)%></u><%=gender(3)%>����</font></p></td>
<%else%>
    <td>&nbsp;</td>
<%end if%>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><p><em><font size="3" face="�з���" >����D�ѹq�������L�o�A�L��ñ�p </font></em><em></em></p></td>
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
