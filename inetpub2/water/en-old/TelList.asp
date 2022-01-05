<!-- #include file="../conn.asp" -->

<%
mMonth = request("mMonth")
 mndate=right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())


if IsNumeric(mMonth) then
	if int(mMonth)<1 or int(mMonth)>12 then
		response.redirect "birthdayList.asp"
	end if
else
	response.redirect "birthdayList.asp"
end if

SQl =   "select m.memno, m.memname, m.memcname, CONVERT(varchar, m.membday, 107) as 'membday', " & _
        "	ISNULL(m.memaddr1, '') AS 'memaddr1', ISNULL(m.memaddr2, '') AS 'memaddr2', " & _
        "	ISNULL(m.memaddr3, '') AS 'memaddr3', " & _
        "	m.mstatus,m.memMobile,m.memEmail  ,m.accode " & _
        "from memmaster m,  (SELECT DISTINCT memNo FROM Share WHERE amount > 0) s " & _
        "WHERE m.memNo = s.memno " & _
        "AND m.mstatus NOT IN ('D', 'V','C','B','P','8','9') " & _
        "AND wdate is null " & _
        "order by m.memno "
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn

if request.form("output")="word" then
	Response.ContentType = "application/msword"
elseif request.form("output")="excel" then
	Response.ContentType = "application/vnd.ms-excel"

end if
%>
<html>
<head>
<title>Birthday List</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
<center>
<font size="4"  face="�з���" >
<table border="0" cellpadding="0" cellspacing="0" >
	<tr height="30" valign="top" align="center"><td colspan="15"><font size="4"  face="�з���" >���ȸp���u�x�W���U��<br>
<%=monthname(mMonth)%>�� �����ͤ�W��C��<br><font size="2"  face="�з���" >��� : <%=mndate%></font></font></td></tr>
        

</table>
<br>
<br>

              
<table border="0" cellpadding="0" cellspacing="0"  width="1150">

	<tr height="15" valign="bottom">
		<td width=70 align="center"><font size="3"  face="�з���" >�����s��</font></td>
                <td width=150 align="center"><font size="3"  face="�з���" >�^��W��</font></td> 
		
		<td width=70 align="center"><font size="3"  face="�з���" >����m�W</font></td>
                <td width=70 align="center"><font size="3"  face="�з���" >�X�ͤ��</font></td>
		<td width=100 align="center"><font size="3"  face="�з���" >���p</font></td>
                <td width=230 align="center"><font size="3"  face="�з���" >�p���q��</font></center></td>
		<td width=230 align="center"><font size="3"  face="�з���" >�q�l</font></center></td>						
                <td width=230 align="center"><font size="3"  face="�з���" >�Ƶ�</font></center></td>
	</tr>
	<tr><td colspan=10><hr></td></tr>
<%
do while not rs.eof
                 select case rs("mstatus") 
                       	case  "L"
                              xstatus =  "�b�b"
                             
                        case  "D" 
                             xstatus =  "�N��"  
		
                        case "V"
                              xstatus = "IVA"
                        
                        case "C"
                              xstatus = "�h��" 
		
                        case  "B"
                              xstatus = "�}��" 
			 
                        case  "P"
                              xstatus = "�h�@"
			       
                        case  "N" 
                              xstatus = "���`"
			    
                        case  "J" 
                              xstatus = "�s��"
                          
                        case  "T" 
                              xstatus = "�w��" 
                           
                        case  "H" 
                               xstatus = "�Ȱ��Ȧ�"
			
                        case  "A"
                               xstatus =  "�۰���b(ALL)"
			
                        case "0"
                              xstatus = "�۰���b(�Ѫ�)"
			 
                        case "1"
                              xstatus = "�۰���b(�Ѫ�,�Q��)"
			
                        case "2"
                              xstatus = "�۰���b(�Ѫ�,����)"                         
			 
                        case "3"
                             xstatus = "�۰���b(�Q��,����)"                         
			 
                        case "M"
                             xstatus = "�w��,�Ȧ�"
			   
                        case "F"
                              xstatus = "�S�O�Ӯ�"  
			  
                        case "8"
                             xstatus = "�פ���y��b"
                          
                        case "9"
                             xstatus = "�פ���y���`"
                          
                     
               end select  
                   idx = ""
                  if rs("accode")="9999" then idx="�h��" end if
                            
%>
	<tr>
		<td width=70 align="center"><%=rs("memNo")%></td>
                <td width=150 align="center"><%=rs("memName")%></td>
		
		<td width=70 align="center"><font size="3"  face="�з���" ><%=rs("memcname")%></font></td>		
		<td width=70 align="center"><%=right("0"&day(rs("memBday")),2)&"/"&right("0"&month(rs("memBday")),2)&"/"&year(rs("memBday"))%></td>
		<td width=100 align="center"><font size="2"  face="�з���" ><%=xstatus%></font></td>
             
                <td width=230 align="left"><font size="2"  face="�з���" ><%=rs("memMobile")%></font></td>
                <td width=230 align="left"><font size="2"  face="�з���" ><%=rs("memEmail")%></font></td>		
                <td width=230 align="left"><font size="2"  face="�з���" ><%=idx%></font></td>

	</tr>
<%
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
