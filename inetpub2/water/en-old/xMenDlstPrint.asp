<!-- #include file="../conn.asp" -->

<%
server.scripttimeout = 1800


xstatus = request.form("KIND")


if xstatus="all" then
      stylefield = ""
else
      stylefield = " mstatus = '"&xstatus&"' "
end if
SQl = "select *  from memmaster  where "&stylefield
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn,1,1

ttlamt = 0
ttlsamt = 0
ttlpamt = 0
ttlpint = 0
ttlisamt = 0
ttlipamt = 0
ttlipint = 0
xdate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
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
	objFile.Write "������b��Ʋө���"
	objFile.WriteLine ""	
	objFile.Write "���"&":"&xdate
	objFile.WriteLine ""	
	objFile.WriteLine ""	
	objFile.Write "�����s��   "
	objFile.Write "         �����W��         "
	objFile.Write "   �Ȧ�b�� "
	objFile.Write "   �C���x�W(�Ȧ�) "
        objFile.Write " ��b�W�� "
        objFile.Write " �C���x�W(�w��) "
	objFile.Write "  �w�йL�� "
	objFile.Write "   �{�p "

	objFile.WriteLine ""
	for idx = 1 to 120
		objFile.Write "-"
	next
	objFile.WriteLine ""

	do while not rs.eof
 
                select case rs("mstatus") 
                       	case  "L"
                              idx =  "�b�b"
                             
                        case  "D" 
                             idx =  "�N��"  
		
                        case "V"
                              idx = "IVA"
                        
                        case "C"
                              idx = "�h��" 
		
                        case  "B"
                              idx = "�h�@" 
			 
                        case  "P"
                              idx = "�}��"
			       
                        case  "N" 
                              idx = "���`"
			    
                        case  "J" 
                              idx = "�s��"
                          
                        case  "T" 
                              idx = "�w��" 
                           
                        case  "H" 
                               idx = "�Ȱ��Ȧ�"
			
                        case  "A"
                               idx =  "�۰���b(ALL)"
			
                        case "0"
                              idx = "�۰���b(�Ѫ�)"
			 
                        case "1"
                              idx = "�۰���b(�Ѫ�,�Q��)"
			
                        case "2"
                              idx = "�۰���b(�Ѫ�,����)"                         
			 
                        case "3"
                             idx = "�۰���b(�Q��,����)"                         
			 
                        case "M"
                             idx = "�w��,�Ȧ�"
			   
                        case "F"
                              idx = "�S�O�Ӯ�"  
			  
                        case "8"
                             idx = "�פ���y��b"
                          
                        case "9"
                             idx = "�פ���y���`"
                          
                     
               end select 
            Bank = rs("bnk")&"-"&rs("bch")&"-"&rs("bacct")
            if Bank = "--"  then
               Bank = ""
            end if
            monthsave = rs("monthsave")
            if monthsave <>""  then
                monthsave = cint(monthsave)
            else 
               monthsave=0                          
            end if  
            bnklmt = rs("bnklmt")
            if bnklmt <>""  then
               bnklmt = cint(bnklmt) 
            else  
               bnklmt=0                          
            end if  
            monthssave = rs("monthssave")
            if monthssave <>""  then
               monthssave = cint(monthssave)
            else 
               monthssave= 0                           
            end if 
            tpayamt = rs("tpayamt")
            if tpayamt <>""   then
               tpayamt = cint(tpayamt) 
            else
               tpayamt=0                           
            end if 

  		objFile.Write left(" "&rs("memNo")&spaces,10) 
		objFile.Write left(rs("memname")&" "&rs("memcname")&spaces,22)
		objFile.Write left(bank&spaces,15)
		objFile.Write right(spaces&formatnumber(monthsave,2),13)
		objFile.Write right(spaces&formatnumber(bnklmt,2),13)
		objFile.Write right(spaces&formatnumber(monthssave,2),13)
		objFile.Write right(spaces&formatnumber(tpayamt,2),13)
		objFile.Write right(spaces&idx,10)              
		objFile.WriteLine    

                
 
	rs.movenext
	loop

	for idx = 1 to 120
		objFile.Write "-"
	next
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
<title>�U�ڱb�Ӷ��C��</title>
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
	<td colspan="15"><font size="4">������b��Ʋө���</font></td>
	</tr>
	<tr height="35" valign="top" align="center">
	<td colspan="15"><font size="4">��� : <%=xdate%></font></td>
        </tr>
</center>
</table>
<table border="0" cellspacing="1" cellpadding="4" align="center" bgcolor="336699">
	<tr bgcolor="#330000" align="center">
		
	<td><font size="2" color="#FFFFFF">�����s��</font></td>
	<td><font size="2" color="#FFFFFF">�����W��</font></td>
	<td><font size="2" color="#FFFFFF">�Ȧ�b��</font></td>
	<td><font size="2" color="#FFFFFF">�C���x�W(�Ȧ�)</font></td>
	<td><font size="2" color="#FFFFFF">��b�W��</font></td>
	<td><font size="2" color="#FFFFFF">�C���x�W(�w��)</font></td>
	<td><font size="2" color="#FFFFFF">�w�йL��</font></td>
	<td><font size="2" color="#FFFFFF">�{�p</font></td>
	</tr>
	
<%
   if not rs.eof then

    do while not rs.eof
        
                select case rs("mstatus") 
                       	case  "L"
                              idx =  "�b�b"
                             
                        case  "D" 
                             idx =  "�N��"  
		
                        case "V"
                              idx = "IVA"
                        
                        case "C"
                              idx = "�h��" 
		
                        case  "B"
                              idx = "�h�@" 
			 
                        case  "P"
                              idx = "�}��"
			       
                        case  "N" 
                              idx = "���`"
			    
                        case  "J" 
                              idx = "�s��"
                          
                        case  "T" 
                              idx = "�w��" 
                           
                        case  "H" 
                               idx = "�Ȱ��Ȧ�"
			
                        case  "A"
                               idx =  "�۰���b(ALL)"
			
                        case "0"
                              idx = "�۰���b(�Ѫ�)"
			 
                        case "1"
                              idx = "�۰���b(�Ѫ�,�Q��)"
			
                        case "2"
                              idx = "�۰���b(�Ѫ�,����)"                         
			 
                        case "3"
                             idx = "�۰���b(�Q��,����)"                         
			 
                        case "M"
                             idx = "�w��,�Ȧ�"
			   
                        case "F"
                              idx = "�S�O�Ӯ�"  
			  
                        case "8"
                             idx = "�פ���y��b"
                          
                        case "9"
                             idx = "�פ���y���`"
                          
                     
               end select 
            Bank = rs("bnk")&"-"&rs("bch")&"-"&rs("bacct")
            if Bank = "--"  then
               Bank = ""
            end if
            monthsave = rs("monthsave")
            if monthsave ="" or monthsave=0  then
               monthsave=""
            end if  
            bnklmt = rs("bnklmt")
            if bnklmt ="" or bnklmt=0  then
               bnklmt=""
            end if  
            monthssave = rs("monthssave")
            if monthssave ="" or monthssave=0  then
               monthssave=""
            end if 
            tpayamt = rs("tpayamt")
            if tpayamt ="" or tpayamt=0  then
               tpayamt=""
            end if 
%>
   <tr bgcolor="#FFFFFF">
	
  	<td><font size="2"><%=rs("memno")%></font></td>
	<td><font size="2"><%=rs("memname")%><%=rs("memcname")%></font></td>
	<td><font size="2"></font><%=Bank%></td>	
	<td align="right"><font size="2"><%=monthsave%></font></td>
	<td align="right"><font size="2"><%=bnklmt%></font></td>
	<td align="right"><font size="2"><%=monthssave%></font></td>
	<td align="right"><font size="2"><%=tpayamt%></font></td>
	<td align="right"><font size="2"><%=idx%></font></td>
   </tr> 
<%	

   rs.movenext
   loop
end if
%>

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
