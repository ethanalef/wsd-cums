<!-- #include file="../conn.asp" -->

<%
server.scripttimeout = 1800

stdate1 = request.form("sdate1")
stdate2 = request.form("sdate2")
yy = right(stdate1,4)
mm = cint(mid(stdate1,4,2))
dd = cint(left(stdate1,2))
todate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
stdate = yy&"."&right("0"&mm,2)&"."&right("0"&dd,2)

yy = right(stdate2,4)
mm = cint(mid(stdate2,4,2))
dd = cint(left(stdate2,2))

eddate = yy&"."&right("0"&mm,2)&"."&right("0"&dd,2)


    stylefield = " and a.ldate >= '"&stdate&"' and a.ldate <= '"&eddate&"' "  
 


     select case request.form("KIND")
            case "cash"
		stylefield =stylefield&" and a.code='A3' order by a.memno "
            case "bank"
                 stylefield =stylefield&" and a.code='A1' order by a.memno "
            case "Trea"
                stylefield =stylefield&" and a.code='A2'  order by a.memno " 

           case "nacct"
                stylefield =stylefield&" and a.code='0A' order by a.memno "

           case "cfee"
                stylefield =stylefield&" and (a.code='G0' or a.code='G3') order by a.memno,a.ldate "
           case "bfee"
                stylefield =stylefield&" and (a.code='H0' or a.code='H3') order by a.memno "
           case "Divid"
                stylefield =stylefield&" and left(a.code,1)='C'  order by a.memno ,a.ldate "
           case "ploan"
                stylefield =stylefield&" and (a.code='B0' ) order by a.memno "
           case "swithd"
                stylefield =stylefield&" and (a.code='B1' )  order by a.memno  "
           case "adj"
                stylefield =stylefield&" and (a.code='A7' )  order by a.memno  "
           case "ins"
                stylefield =stylefield&" and (a.code='A4' )  order by a.memno  "
           case "all"
                stylefield =stylefield&" order by a.memno "
     end select


SQl = "select a.memno,a.code,a.ldate,a.amount,b.memname,b.memcname,convert(char(10),ldate,102) as xdate  from share a,memmaster b where a.memno=b.memno  "&stylefield
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn


ttlamt = 0
ttlsamt = 0
ttlpamt = 0
ttlpint = 0
ttlisamt = 0
ttlipamt = 0
ttlipint = 0

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
	objFile.Write "�Ѫ��b�Ӷ��ө���"
	objFile.WriteLine ""	
	objFile.Write "���"&":"&right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
	objFile.WriteLine ""	
	objFile.WriteLine ""	
	objFile.Write " �����s�� "
	objFile.Write "         �����W��                 "
	objFile.Write "            ������   "
        objFile.Write "    ���O   "
	objFile.Write right(spaces&"���B",16)
	objFile.WriteLine ""
	for idx = 1 to 101
		objFile.Write "-"
	next
	objFile.WriteLine ""
        xmemno =rs("memno") 

	do while not rs.eof

   
             xdate =right("0"&day(rs("ldate")),2)&"/"&right("0"&month(rs("ldate")),2)&"/"&year(rs("ldate"))
              select case rs("code")

		     case "0A"
                          skind = "�s  ��"  
                          samt = rs("amount")
		     case "A1"
                          skind = "��  ��"  
                          samt = rs("amount")
                     case "A2"
			  skind = "�w  ��" 
                          samt =rs("amount")
                     case "A3"
			  skind = "�{  ��" 
                          samt =rs("amount")
                     case "A4"
			  skind = "�O�I��" 
                          samt =rs("amount")
                     case "A7"
			  skind = "��  ��" 
                          samt =rs("amount")
                     case "C0","C1","C2","C3"
			  skind = "��  ��" 
                          samt =rs("amount")            	
                     case "H0","H3"
			  skind = "��|�O" 
                          samt =rs("amount")  
                     case "G0","G3"
			  skind = "�|  �O" 
                          samt =rs("amount")                        
                     case "B0"
                           skind="�h���ٶU��"
                           samt = rs("amount")
                     case "B1"
                            skind="�h  ��"
                           samt = rs("amount")                                                                
                  
               end select   
                 
                name1 = left(rs("memname")&spaces,24)
                name2 = left(rs("memcname")&spaces,10)   
                xkind = left(skind&spaces,10)  
                ttlname = name1&name2
                ttlTemp = ttlTemp +samt   
  		objFile.Write left(" "&xmemNo&spaces,10) 
		objFile.Write left(ttlname&spaces,36)
                objFile.Write left("     "&xdate&spaces,22)
		objFile.Write left(xkind&spaces,12)
		objFile.Write right(spaces&formatnumber(samt,2),13)
		objFile.WriteLine    


                xmemno = rs("memno")
                sttlamt = 0
                pamt1 = 0
                pint1 = 0
                samt = 0
                pint2 = 0
                pamt2 = 0
   
            
        
 
             
	rs.movenext
	loop
	for idx = 1 to 101
		objFile.Write "-"
	next
	objFile.WriteLine ""
	objFile.Write space(72)
	objFile.Write  left("�`���B �G"&spaces,10)
	objFile.Write right(spaces&formatnumber(ttlTemp,2),13)
	objFile.WriteLine ""

	objFile.Close

	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.redirect "../txt/"&session("username")&".txt"
end if
%>
<html>
<head>
<title>�Ѫ��b�Ӷ��C��</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
<center>
<table border="0" cellpadding="0" cellspacing="0" >
	<tr height="30" valign="top" align="center">
	<td colspan="15"><font size="4">���ȸp���u�x�W���U��</font></td>
	</tr>
	<tr height="35" valign="top" align="center">
	<td colspan="15"><font size="4">�Ѫ��b�Ӷ��ө���</font></td>
	</tr>
	<tr height="35" valign="top" align="center">
	<td colspan="15"><font size="4">��� : <%=todate%></font></td>
        </tr>
</center>
</table>
<table border="0" cellspacing="1" cellpadding="4" align="center" bgcolor="336699">
	<tr bgcolor="#330000" align="center">
		
	<td><font size="2" color="#FFFFFF">�����s��</font></td>
	<td><font size="2" color="#FFFFFF">�����W��</font></td>
	<td><font size="2" color="#FFFFFF"> ��� </font></td>
	<td><font size="2" color="#FFFFFF"> ���O </font></td>
	<td><font size="2" color="#FFFFFF"> ���B </font></td>
	</tr>
	
<%
   if not rs.eof then
        xmemno =rs("memno") 

	do while not rs.eof
 

                   
             xdate =right("0"&day(rs("ldate")),2)&"/"&right("0"&month(rs("ldate")),2)&"/"&year(rs("ldate"))
              select case rs("code")

		     case "0A"
                          skind = "�s  ��"  
                          samt = rs("amount")
		     case "A1"
                          skind = "��  ��"  
                          samt = rs("amount")
                     case "A2"
			  skind = "�w  ��" 
                          samt =rs("amount")
                     case "A3"
			  skind = "�{  ��" 
                          samt =rs("amount")
                     case "A4"
			  skind = "�O�I��" 
                          samt =rs("amount")
                     case "A7"
			  skind = "��  ��" 
                          samt =rs("amount")
                     case "C0"
			  skind = "��  ��" 
                          samt =rs("amount")            	
                     case "C1"

			  skind = "�Ѯ��L�b�ܻȦ�" 
                          samt =rs("amount")     
                     case "H0","H3"
			  skind = "��|�O" 
                          samt =rs("amount")  
                     case "G0","G3"
			  skind = "�|  �O" 
                          samt =rs("amount")                        
                     case "B0"
                           skind="�h���ٶU��"
                           samt = rs("amount")
                     case "B1"
                            skind="�h  ��"
                           samt = rs("amount")   
                     case "BE"                                                             
			  skind = "�Ѱy�٥�" 
                          samt =rs("amount") 
                     case "BF"                                                             
			  skind = "�Ѱy�ٮ�" 
                          samt =rs("amount")                        
               end select 
              ttlsamt = ttlsamt + samt

%>
   <tr bgcolor="#FFFFFF">
	
  	<td><font size="2"><%=rs("memno")%></font></td>
	<td><font size="2"><%=rs("memname")%><%=rs("memcname")%></font></td>
	<td ><font size="2"><%=xdate%></font></td>
	<td align="right"><font size="2" ><center><%=skind%></center></font></td>
	<td align="right"><font size="2"><%=formatnumber(samt,2)%></font></td>
   </tr> 
<%	
                xmemno = rs("memno")
                samt = 0
                                                               
rs.movenext
loop
end if
%>
	<tr>
		<td></td>
		<td></td>              		
                 <td></td>
		 <td>�`���B �G</td>
                
	
		<td width=100 align="right"><%=formatNumber(ttlsamt,2)%></td>
	

	</tr>
</table>
</center>
</body>
</html>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
