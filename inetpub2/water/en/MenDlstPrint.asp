<!-- #include file="../conn.asp" -->

<%
server.scripttimeout = 1800

 mndate=right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
xstatus = request.form("KIND")
banks   = request.form("banks")
bnklmt  = request.form("banklmt")
tpay    = request.form("tpay")

   stylefield = ""
if xstatus="all" then

   stylefield = ""
else


  select case xstatus 
         case "A"

              if banks = "Y" and bnklmt="Y" then

                 stylefield = " where (mstatus='A' ) and (monthsave=0 or monthsave is null ) and (bnklmt = 0 or bnklmt is null )  "
                 
              else
              if banks = "Y" and bnklmt<>"Y"  then
                 stylefield = " where (mstatus='A' ) and (monthsave=0 or monthsave is null )  "
              else

             if bnklmt="Y" and banks<>"Y"  then
                stylefield = " where (mstatus='A') and (bnklmt = 0 or bnklmt is null ) " 
             else
                stylefield = " where ( mstatus='A' ) "
            end if  
            end if
            end if
           case "0"
             if banks = "Y" and bnklmt="Y" then

                 stylefield = " where ( mstatus='0' ) and (monthsave=0 or monthsave is null ) and (bnklmt = 0 or bnklmt is null )  "
                 
              else
              if banks = "Y" and bnklmt<>"Y"  then
                 stylefield = " where ( mstatus='0' ) and (monthsave=0 or monthsave is null )  "
              else

             if bnklmt="Y" and banks<>"Y"  then
                stylefield = " where ( mstatus='0' ) and (bnklmt = 0 or bnklmt is null ) " 
            else
                stylefield = " where ( mstatus='0' ) "
            end if  
            end if
            end if
          case "1"
             if banks = "Y" and bnklmt="Y" then

                 stylefield = " where ( mstatus='1' ) and (monthsave=0 or monthsave is null ) and (bnklmt = 0 or bnklmt is null )  "
                 
              else
              if banks = "Y" and bnklmt<>"Y"  then
                 stylefield = " where ( mstatus='1' ) and (monthsave=0 or monthsave is null )  "
              else

             if bnklmt="Y" and banks<>"Y"  then
                stylefield = " where (  mstatus='1') and (bnklmt = 0 or bnklmt is null ) " 
             else
                stylefield = " where ( mstatus='1' ) "
            end if  
            end if
            end if
         case "T","M"
            if tpay = "Y" then
                 stylefield = " where (mstatus='T'  or mstatus='M') and (bnklmt = 0 or bnklmt is null ) " 
            else
                 stylefield = " where (mstatus='T'  or mstatus='M') "
            end if                  
        case else
            stylefield = "where  mstatus = '"&xstatus&"' "
end select
end if


SQl = "select *  from memmaster "&stylefield&" and wdate is null  order by memno"
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
		
                        case  "P"
                              idx = "�h�@" 
			 
                        case  "B"
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
	<tr height="30" valign="top" align="center"><td colspan="15"><font size="4"  face="�з���" >���ȸp���u�x�W���U��<br>������b��Ʋө���<br><font size="2"  face="�з���" >��� : <%=mndate%></font></font></td></tr>
        

</table>
<br>
<br>
<table border="0" cellpadding="0" cellspacing="0">
	<tr height="20" valign="bottom">
	<td width=70 align="center"><font size="3"  face="�з���" >�����s��</font></td>
	<td width=70 align="center"><font size="3"  face="�з���" >�����W��</font></td>
	<td width=130 align="center"><font size="3"  face="�з���" >�Ȧ�b��</font></td>
	<td width=100 align="center"><font size="3"  face="�з���" >�C���x�W<br>(�Ȧ�)</font></td>
	<td width=100 align="center"><font size="3"  face="�з���" >��b�W��</font></td>
	<td width=100 align="center"><font size="3"  face="�з���" >�C���x�W<br>(�w��)</font></td>
	<td width=100 align="center"><font size="3"  face="�з���" >�w�йL��</font></td>
	<td width=130 align="center"><font size="3"  face="�з���" >�{�p</font></td>
	</tr>
	<tr><td colspan=8><hr></td></tr>
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
		
                        case  "P"
                              idx = "�h�@" 
			 
                        case  "B"
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
	
  	<td width=70 align="center"><%=rs("memno")%></td>
	<td width=70 align="center"><font size="3"  face="�з���" ><%=rs("memcname")%></font></td>
	<td width=130 align="center"><%=Bank%></td>	
	<td width=100 align="center"><%=monthsave%></td>
	<td width=100 align="center"><%=bnklmt%></td>
	<td width=100 align="center"><%=monthssave%></td>
	<td width=100 align="center"><%=tpayamt%></td>
	<td width=130 align="center"><font size="3"  face="�з���" ><%=idx%></font></td>
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

