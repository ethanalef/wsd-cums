<!-- #include file="../conn.asp" -->

<%
server.scripttimeout = 1800

stdate1 = request.form("stdate1")
stdate2 = request.form("stdate2")
yy = right(stdate1,4)
mm = mid(stdate1,4,2)
dd = left(stdate1,2)

todate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())

stdate1 = yy&"/"&mm&"/"&dd

yy = right(stdate2,4)
mm = mid(stdate2,4,2)
dd = left(stdate2,2)

stdate2 = yy&"/"&mm&"/"&dd



stylefield="and (Appdate>='"&stdate1&"' and Appdate<='"&stdate2&"' )  "

memno=request.form("memNo")
if memno <>"*" then
   stylefield =stylefield&" and a.memno='"&memno&"' " 
end if
    kind = request.form("KIND")
     select case request.form("KIND")
            case "CA"
		stylefield =stylefield&" and FirstApproval ='Approved' order by uid "
            case "CR"
                 stylefield =stylefield&" and FirstApproval ='Rejected' order by uid "
            case "CP"
          	stylefield =stylefield&" and FirstApproval ='Pending' order by uid "
            case "DA"
		stylefield =stylefield&" and SecondApproval ='Approved' order by uid "
            case "DR"
                 stylefield =stylefield&" and SecondApproval ='Rejected' order by uid "
            case "DP"
          	stylefield =stylefield&" and SecondApproval ='Pending' order by uid "
            case "all"

 
     end select


SQl = "select a.uid,a.memno,a.appdate,a.loanamt,a.installment,a.chequeamt,a.firstapproval,a.secondapproval,a.guarantorID,a.guarantor2ID,a.guarantor3ID,b.memname,b.memcname  from loanapp a,memmaster b where a.memno=b.memno "&stylefield
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn



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
	objFile.Write "�U�ڥӽвө���"
	objFile.WriteLine ""	
	objFile.Write "���"&":"&right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
	objFile.WriteLine ""	
	objFile.WriteLine ""	
	objFile.Write " �ӽЧǸ�"
	objFile.Write "   ���"
	objFile.Write "    �����s�� "
	objFile.Write "      �����W��            "
        objFile.Write "     �U���`�B "
	objFile.Write " ����"
	objFile.Write " �C���ٴ�"
	objFile.Write "   �������B   "
	objFile.Write "  ���p  "
	objFile.WriteLine ""
	for idx = 1 to 110
		objFile.Write "-"
	next
	objFile.WriteLine ""
        ttlcnt = 0
     	ttlapamt = 0
        ttlcchqamt = 0      
     
	do while not rs.eof
                xdate = right("0"&day(rs("appdate")),2)&"/"&right("0"&month(rs("appdate")),2)&"/"&year(rs("appdate"))      
               monthrepay = rs("loanamt")/rs("installment")
               if int(monthrepay) = monthrepay then
                  monthrepay= int(monthrepay)
               else
                  monthrepay = int(monthrepay)+1
               end if
               ttlapamt = ttlapamt +rs("loanamt")
               ttlchqamt = ttlchqamt + rs("chequeamt")
               ttlcnt = ttlcnt + 1
               if rs("firstApproval")<>"" then
               select case rs("firstApproval")
                      case  "Approved"
                             status ="���ӽ�"
                      case  "Rejected"
                            status = "�ڵ��ӽ�" 
                      case "Pending"
                           status = "�f�z��"
              end select
             end if
             if rs("SecondApproval")<>"" then
              select case rs("SecondApproval")
                      case  "Approved"
                             status ="���Ƨ��ӽ�"
                      case  "Rejected"
                            status = "���Ʃڵ��ӽ�" 
                      case "Pending"
                           status = "���Ƽf�z��"
              end select
            end if
 
               ttlname  = rs("memname")&" "&rs("memcname")
                objFile.Write right(spaces&rs("uid"),8) 
                objFile.Write left(" "&xdate&spaces,11)
  		objFile.Write left("  "&rs("memNo")&spaces,8) 
		objFile.Write left("  "&ttlname&spaces,25)                
		objFile.Write right(spaces&formatnumber(rs("loanamt"),2),13)
		objFile.Write right(spaces&formatnumber(rs("installment"),0),4)
		objFile.Write right(spaces&formatnumber(monthrepay,2),10)
                objFile.Write right(spaces&formatnumber(rs("chequeamt"),2),13)
            
		objFile.Write right(spaces&status,8)
		objFile.WriteLine    

                ttlcnt = ttlcnt + 1
 
                                 
           

             
	rs.movenext
	loop
	for idx = 1 to 110
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
<title>�U�ڱb�C��</title>
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
	<td colspan="15"><font size="4">�U�ڥӽвө���</font></td>
	</tr>
	<tr height="35" valign="top" align="center">
	<td colspan="15"><font size="4">��� : <%=todate%></font></td>
        </tr>
</center>
</table>
<table border="0" cellspacing="1" cellpadding="4" align="center" bgcolor="336699">
	<tr bgcolor="#330000" align="center">
	<td><font size="2" color="#FFFFFF">�ӽЧǸ�</font></td>
	<td><font size="2" color="#FFFFFF">�ӽФ��</font></td>		
	<td><font size="2" color="#FFFFFF">�����s��</font></td>
	<td><font size="2" color="#FFFFFF">�����W��</font></td>
	<td><font size="2" color="#FFFFFF">�U���`�B</font></td>
	<td><font size="2" color="#FFFFFF">����</font></td>
	<td><font size="2" color="#FFFFFF">�C���ٴ�</font></td>
	<td><font size="2" color="#FFFFFF">�������B</font></td>

	<td><font size="2" color="#FFFFFF">�{��</font></td>
	
	</tr>
	
<%
        ttlcnt = 0
     	ttlapamt = 0
        ttlcchqamt = 0      
     
	do while not rs.eof
                xdate = right("0"&day(rs("appdate")),2)&"/"&right("0"&month(rs("appdate")),2)&"/"&year(rs("appdate"))      
               monthrepay = rs("loanamt")/rs("installment")
               if int(monthrepay) = monthrepay then
                  monthrepay= int(monthrepay)
               else
                  monthrepay = int(monthrepay)+1
               end if
               ttlapamt = ttlapamt +rs("loanamt")
               ttlchqamt = ttlchqamt + rs("chequeamt")
               ttlcnt = ttlcnt + 1
               if rs("firstApproval")<>"" then
               select case rs("firstApproval")
                      case  "Approved"
                             status ="���ӽ�"
                      case  "Rejected"
                            status = "�ڵ��ӽ�" 
                      case "Pending"
                           status = "�f�z��"
              end select
             end if
             if rs("SecondApproval")<>"" then
              select case rs("SecondApproval")
                      case  "Approved"
                             status ="���Ƨ��ӽ�"
                      case  "Rejected"
                            status = "���Ʃڵ��ӽ�" 
                      case "Pending"
                           status = "���Ƽf�z��"
              end select
            end if
%>
   <tr bgcolor="#FFFFFF">
        <td><font size="2"><%=rs("uid")%></font></td>
	<td ><font size="2"><%=xdate%></font></td>	
  	<td><font size="2"><%=rs("memno")%></font></td>
        <td><font size="2"><%=rs("memname")%><%=rs("memcname")%></font></td>       
	<td align="right"><font size="2" ><%=formatnumber(rs("loanamt"),2)%></font></td>
	<td align="center"><font size="2"><%=formatnumber(rs("installment"),0)%></font></td>
	<td align="right"><font size="2"><%=formatnumber(monthrepay,2)%></font></td>
	<td align="right"><font size="2"><%=formatnumber(rs("chequeamt"),2)%></font></td>
                 
 	<td align="center"><font size="2"><%=status%></font></td>
   </tr> 
<%	

rs.movenext
loop
%>
	<tr>
		<td>�`��:</td>
		<td><%=ttlcnt%></td>              		                 
		 <td></td>
                <td>���B �G</td>
		<td width=100 align="right"><font size="2"><b><%=formatNumber(ttlapamt,2)%></b></font></td>
		<td></td>		
                <td></td>
                <td width=100 align="right"><font size="2"><b><%=formatNumber(ttlchqamt,2)%></b></font></td>
	

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
