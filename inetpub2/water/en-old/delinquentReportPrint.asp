<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<%

mndate=right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())

   
   nday = request.form("noofday")
   YY = YEAR(DATE())
   MM = MONTH(DATE())
   dd = DAY(DATE())
  
  chkdate =DATESERIAL(YY,MM,DD-NDAY)
 
SQl = "SELECT  a.memno, a.lnnum, MAX(a.ldate) AS maxdate, b.appamt, b.bal, c.memname, "&_
      "c.memCName,C.mstatus FROM loan a INNER JOIN Loanrec b ON a.memno = b.memno AND a.lnnum = b.lnnum INNER JOIN "&_
      "MemMaster c ON a.memno = c.memNo WHERE (b.repaystat = 'N') AND (c.Wdate IS NULL) "&_
      "GROUP BY  a.memno, a.lnnum, c.memname, c.memCName, b.appamt, b.bal,c.mstatus "
      Set rs = Server.CreateObject("ADODB.Recordset") 
      rs.open sql, conn, 3


if request.form("output")="word" then
	Response.ContentType = "application/msword"
elseif request.form("output")="excel" then
	Response.ContentType = "application/vnd.ms-excel"

end if
%>
<html>
<head>
<title>Delinquent Report</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<LINK href="../main.css" rel=STYLESHEET type=text/css>
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
<center>
<table border="0" cellpadding="0" cellspacing="0" >
	<tr height="30" valign="top" align="center"><td colspan="15"><font size="4"  face="�з���" >���ȸp���u�x�W���U��<br>�b�b�C��<br><font size="2"  face="�з���" >��� : <%=mndate%></font></font></td></tr>
        <tr height="30" ><td colspan=9></td></tr>


<table border="0" cellpadding="0" cellspacing="0">
	<tr height="15" valign="bottom">
		<td width=70 align="center"><font size="2"  face="�з���" >�����W��</font></td>
               <td width=180 align="center"><font size="3"  face="�з���" >�^��m�W</font></td>
		<td width=70 align="center"><font size="3"  face="�з���" >����m�W</font></td>
		<td width="70" align="center"><font size="2"  face="�з���" >�U�ڽs��</font></td>
		<td width="130" align="right"><font size="2"  face="�з���" >�U���`�B</font></td>
		<td width="130" align="right"><font size="2"  face="�з���" >�U�ڵ���</font></td>
                <td width="80" align="right"><font size="2"  face="�з���" >���p</font></td>
	</tr>
	<tr><td colspan=7><hr></td></tr>
<%

do while not rs.eof
   if rs("maxdate") <= chkdate then
                   select case rs("mstatus")
                       case "L"
                           xstatus= "�b�b"
                       case  "D"
                           xstatus="�N��"
                       
                       case  "V"
                           xstatus= " IVA "
                         
                       case  "C"
                             xstatus= "�h��"
             
                       case  "P"
                             xstatus= "�h�@"
                         
                       case  "B"
                            xstatus="�}��"
                    
                       case  "N"
                            xstatus= "���`"
                        
                      case  "J"
                            xstatus= "�s��"
                       
                      case "H"
                          xstatus= "�Ȱ��Ȧ�"
                      
                       case  "A"
                            xstatus="�۰���b"

                       case  "0"
                            xstatus="�۰���b(�Ѫ�)"                       
                       case  "1"
                            xstatus="�۰���b(�Ѫ�,�Q��)"
                       case  "Z"
                            xstatus="�۰���b(�Ѫ�,����)"
                       case "3"
                             xstatus="�۰���b(�Q��,����)"
                       case  "M"
                           xstatus = "�w��,�Ȧ�"
                      
                      case  "T"
                            xstatus= "�w��"
                     case "F"
                          xstatus =  "���D�U��"
                end select

%>
	<tr>
		<td width=70 align="center"><%=rs("memNo")%></td>
                <td width=180 align="left"><%=rs("memname")%></td> 
		<td width=70 align="center"><font size="3"  face="�з���" ><%=rs("memcname")%></font> </td>
		<td width=70 align="center"><%=rs("lnnum")%></td>
		<td width="130" align="right"><%=formatNumber(rs("appamt"),2)%></td>
		<td width="130" align="right"><%=formatNumber(rs("bal"),2)%></td>
                <td width="80" align="center"><font size="2"  face="�з���" ><%=xstatus%></td>
	</tr>
<%
  end if
	rs.movenext
    loop
%>
</table>
</font>
</center>
</body>
</html>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
