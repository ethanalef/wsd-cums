<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<%

mPeriod = request.form("mPeriod")
rate    = request.form("rate")

mndate=right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
myear = year(date())

         set rs = conn.execute("select a.memno,a.dividend,a.bank,b.memname,b.memcname from dividend a,memmaster b where a.memno=b.memNo order by 1  ")
 



ttlamt = 0
    
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
	objFile.Write "���ȸp���u�x�W���U��"
	objFile.WriteLine ""
	objFile.Write  left(space2,21)
	objFile.Write "�Ѯ��C�� - "
        objFile.write myear
	objFile.WriteLine ""	
        objFile.WriteLine ""	
        objFile.WriteLine ""	
	objFile.Write left(spaces,10)
	objFile.Write left("    ����"&spaces,15)
        objFile.Write left("    �m�W"&spaces,40)
	objFile.Write left("    ���B"&spaces,15)
        objFile.Write left("    ����"&spaces,15)
	objFile.WriteLine ""
	for idx = 1 to 100
		objFile.Write "-"
	next       
	objFile.WriteLine ""   
 ttlamt = 0
   shamt = 0
   bkamt = 0
   chamt = 0
   ttlcnt = 0
   shcnt = 0
   bkcnt = 0
   chcnt = 0 
 do while not rs.eof 
      select case rs("bank")
             case "S"
                  idx = "�Ѫ�"
                  shamt = shamt + rs(1)
                  shcnt = shcnt + 1
             case "B"
                  idx="�Ȧ���b"
                  bkamt = bkamt + rs(1) 
                  bkcnt = bkcnt + 1
             case "C"
                  idx="�䲼" 
                  chamt = chamt + rs(1)
                  chcnt = chcnt + 1
      end select
	objFile.Write left(spaces,10)
	objFile.Write right(spaces&rs(0)&"    ",12)
        objFile.Write left(rs(4)&spaces,40)
	objFile.Write right(spaces&formatnumber(rs(1),2),15)
        objFile.Write right(spaces&idx,10)
	objFile.WriteLine ""
        ttlamt = ttlamt + rs(1)    
    rs.movenext
    loop
	for idx = 1 to 100
		objFile.Write "-"
	next       
	objFile.WriteLine ""      
	objFile.Write left(spaces,10)
	objFile.Write right(spaces&"    ",18)
        objFile.Write left("   "&spaces,40)
	objFile.Write right(spaces&formatnumber(ttlamt,2),15)
	objFile.WriteLine ""	 
	objFile.Write left(spaces,10)
	objFile.Write right(spaces&"    ",18)
        objFile.Write left("   "&spaces,50)
	objFile.Write right(spaces&"=============",15)
	objFile.WriteLine ""
        objFile.WriteLine "" 
        objFile.Write  "�Ѫ����B�X�@ : "
        objFile.Write   right(spaces&formatNumber(shamt,2),15)
        objFile.Write  "     �Ѫ��H�ƦX�@ : "
        objFile.Write  right(spaces&formatNumber(shcnt,0),15) 
        objFile.WriteLine ""
        objFile.Write  "�Ȧ���b���B�X�@ : "
        objFile.Write   right(spaces&formatNumber(bkamt,2),15)
        objFile.Write  "     �Ȧ���b�H�ƦX�@ : "
        objFile.Write  right(spaces&formatNumber(bkcnt,0),15) 
        objFile.WriteLine ""
        objFile.Write  "�䲼���B�X�@ : "
        objFile.Write   right(spaces&formatNumber(chamt,2),15)
        objFile.Write  "     �䲼�H�ƦX�@ : "
        objFile.Write  right(spaces&formatNumber(chcnt,0),15) 
        objFile.WriteLine "" 
	objFile.Close


	
	
	response.redirect "../txt/"&session("username")&".txt"
end if
%>
<html>
<head>
<title>�Ѯ��C��</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
<center>
<table border="0" cellpadding="0" cellspacing="0">
	<tr height="30" valign="top" align="center"><td colspan="15"><font size="4"  face="�з���" >���ȸp���u�x�W���U��<br>�Ѯ��C��<br><font size="2"  face="�з���" >��� : <%=mndate%></font></font></td></tr>
        <tr height="30" ><td colspan=9></td></tr>


	<tr height="15" valign="bottom">
        
	<td width="80" align="center"><font size="2"  face="�з���" >�����s��</font></td>
	<td width="80"  align="center"><font size="2"  face="�з���" >  �m�W</font</td>
	<td width="130" align="right"><font size="2"  face="�з���" > ���B</fot></td>
        <td width="80" align="center"><font size="2"  face="�з���" > ����</font></td> 
	</tr>
	<tr><td colspan=6><hr></td></tr>
<% ttlamt = 0
   shamt = 0
   bkamt = 0
   chamt = 0
   ttlcnt = 0
   shcnt = 0
   bkcnt = 0
   chcnt = 0 
  
   do while not rs.eof
      select case rs("bank")
             case "S"
                  idx = "�Ѫ�"
                  shamt = shamt + rs(1)
                  shcnt = shcnt + 1
             case "B"
                  idx="�Ȧ���b"
                  bkamt = bkamt + rs(1) 
                  bkcnt = bkcnt + 1
             case "C"
                  idx="�䲼" 
                  chamt = chamt + rs(1)
                  chcnt = chcnt + 1
      end select
      ttlcnt = ttlcnt + 1
     
%>
     <tr>
          <td width="80" align="center"><%=rs(0)%></td>
          <td width="80" align="center" ><font size="2"  face="�з���" ><%=rs(4)%></font></td>
          <td width="130" align="right"><%=formatnumber(rs(1),2)%></td>
          <td width="80" align="center"><font size="2"  face="�з���" ><%=idx%> </font></td>
     </tr>

<%
    ttlamt = ttlamt + rs(1)
     rs.movenext
    loop
%>


	<tr><td colspan=4><hr></td></tr>
        <tr><td></td>
             <td></td>             
             <td width="130" align="right"><%=formatnumber(ttlamt,2)%></td>
              
         </tr>
        <tr><td></td>
             <td></td>             
             <td width="130" align="right">==========</td>
              
         </tr>	


</table>
<BR>
<BR>

<table border="0" cellpadding="0" cellspacing="0">
<tr>
      <td width="200" ><font size="2"  face="�з���" > �Ѫ����B�X�@ :</font></td>
      <td width="100" align="right"><%=formatNumber(shamt,2)%></td>
      <td width="30">
      <td width="150" ><font size="2"  face="�з���" > �Ѫ��H�ƦX�@ :</font></td>
      <td width="100" align="right"><%=formatNumber(shcnt,0)%></td>      
</tr>
 <tr>
      <td width="200" ><font size="2"  face="�з���" > �Ȧ���b���B�X�@ :</font></td>
      <td width="100" align="right"><%=formatNumber(bkamt,2)%></td>
      <td width="30">
      <td width="150" ><font size="2"  face="�з���" > �Ȧ���b�H�ƦX�@ :</font></td>
      <td width="100" align="right"><%=formatNumber(bkcnt,0)%></td>      
</tr>
 <tr>
      <td width="200" ><font size="2"  face="�з���" > �䲼���B�X�@ :</font></td>
      <td width="100" align="right"><%=formatNumber(chamt,2)%></td>
      <td width="30">
      <td width="150" ><font size="2"  face="�з���" > �䲼�H�ƦX�@ :</font></td>
      <td width="100" align="right"><%=formatNumber(chcnt,0)%></td>      
</tr>

	<tr><td colspan=5><hr></td></tr>
        <tr><td></td>
            <td width=100 align="right"><%=formatnumber(ttlamt,2)%></font></td> 
	    <td></td>
            <td></td> 
            <td width=100 align="right"><%=formatnumber(ttlcnt,0)%></font></td>
            <td></td>
        </tr>
        <tr>
            <td width=200 align="right"></td>   
            <td width=100 align="right">==========</td>   
               <td width="30">
            <td width=150 align="right"></td>   
             <td width=100 align="right">==========</td>    
       
          
        </tr>

</table>
</center>
</body>
</html>

