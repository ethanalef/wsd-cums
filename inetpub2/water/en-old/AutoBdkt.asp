<!-- #include file="../conn.asp" -->

<!-- #include file="navigator.asp" -->
<%
if request("del")<>"" then
 
  uid = request("del")

  pos = instr(uid,"?")
  pos1 = instr(uid,"*")
  dif  = pos1 - pos-1
  smemno = left(uid,pos-1) 
  status = mid(uid,pos+1,dif)
  cname  = mid(uid,pos1+1,10)
  select case cname

               case "�Ȧ楻��"
                    ccode ="E1" 
               case "�w�Х���"
                    ccode ="E2"
               case "�Ȧ�Q��" 
                    ccode="F1"   
               case "�w�ЧQ��"
                    ccode="F2"   
               case "�Ȧ�Ѫ�"  
                    ccode="A1" 
               case "�w�ЪѪ�"  
                    ccode="A2"
  end select

  if status="�_" then
     
     conn.execute("update autopay set flag='F'  where memno='"&smemno&"' and right(code,1)='2'  " )
  else
     conn.execute("update autopay set flag=''  where memno='"&smemno&"'  and right(code,1)='2'  " )
  end if   
  memno = smemno
end if
For Each Field in Request.Form
	TheString = Field & "= Request.Form(""" & Field & """)"
	Execute(TheString)
Next
For Each Field in Request.querystring
	TheString = Field & "= Request.querystring(""" & Field & """)"
	Execute(TheString)
Next

if memNo <> "" then
	sql_filter = sql_filter & " and a.memNo = '"&memNo&"' "
end if
if memName <> "" then
	sql_filter = sql_filter & " and b.memName like '%"&memName&"%'"
end if
if memhkid <> "" then
	sql_filter = sql_filter & " and b.memhkid >= '"&memhkid&"'"
end if
select case  sflag
    case "N" 
    sql_filter = sql_filter & " and a.flag <> 'F' "
   case "Y"
    sql_filter = sql_filter & " and a.flag = 'F' "
end select

IF REQUEST("NPAGE") <> "" OR REQUEST("UPAGE") <>"" THEN
   SQL_FILTER  =SESSION("STRSQL")
END IF


sql = "select a.*,b.memname,b.memhkid from autopay a,memmaster b where  a.memno=b.memno and (right(a.code,1)='2' ) "& sql_filter & " order by a.memno"
set rs = server.createobject("ADODB.Recordset")

rs.open sql, conn, 3


if not rs.eof then
        memno= ""
        memhkid = ""
        memname =""
        sflag = ""

	if request("page") <> "" then
		pageno = cint(request("page"))
	else
		pageno = 1
	end if
	rs.pagesize = 10
	pagesize=rs.pagesize
	rs.absolutepage = pageno
	recordcount=rs.recordcount
	pagecount = rs.pagecount
end if
%>
<html>
<head>
<title>�s�U�ڼƫإ�</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
<!--
function formatDate(dateform){
  cDate = dateform.value;
  dSize = cDate.length;
  if (dSize!=0){
    sCount= 0;
    for(var i=0; i < dSize; i++)
      (cDate.substr(i,1) == "/") ? sCount++ : sCount;
    if (sCount == 2){
		ySize = cDate.substring(cDate.lastIndexOf("/")+1,dSize).length;
		if (ySize<2 || ySize>4 || ySize == 3){
		  return false;
		 }
		idxBarI = cDate.indexOf("/");
		idxBarII = cDate.lastIndexOf("/");
		strD = cDate.substring(0,idxBarI);
		strM = cDate.substring(idxBarI+1,idxBarII);
		strY = cDate.substring(idxBarII+1,dSize);
		strM = (strM.length < 2 ? '0'+strM : strM);
		strD = (strD.length < 2 ? '0'+strD : strD);
		if(strY.length == 2)
		  strY = (strY > 50  ? '19'+strY : '20'+strY);
    }else{
    	if (dSize != 8)
			return false;
		strD = cDate.substring(0,2);
		strM = cDate.substring(2,4);
		strY = cDate.substring(4,8);
    }
    dateform.value = strD+'/'+strM+'/'+strY;
    if (!valDate(strM, strD, strY))
      return false;
    else
      return true;
  }
}
//  -->
</script>
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0">
<!-- #include file="menu.asp" -->
<br>
<center>
<%if msg<>"" then response.write "<font color=red>"&msg&"</font>" end if%>
<form name="form1" method="post" action="AutoBdkt.asp">

�����Ҹ��X : <input type="text" name="memhkid" value="<%=memhkid%>" size="10">
�����s�� : <input type="text" name="memNo" value="<%=memNo%>" size="6">
�����m�W : <input type="text" name="memName" value="<%=memName%>" size="10">
���(Y/N) : <input type="text" name="sflag" value="<%=sflag%>" size="1">
<input type="submit" value="�j�M"  onclick="return validating()" class="sbttn">
</form>
<%if recordcount>pagesize then navigator("AutoBdkt.asp?lnnum="&lnnum&"&memhkid="&memhkid&"&memNo="&memNo&"&memName="&memName&"&sflag="&sflag) end if%>
<table border="0" cellspacing="1" cellpadding="4" align="center" bgcolor="336699">
	<tr bgcolor="#330000" align="center">
	
		<td><font size="2" color="#FFFFFF">�����Ҹ��X</font></td>
		<td><font size="2" color="#FFFFFF">�����s��</font></td>
		<td><font size="2" color="#FFFFFF">�����m�W</font></td>
                <td><font size="2" color="#FFFFFF">�b��</font></td>
		<td><font size="2" color="#FFFFFF">��b���B</font></td>
                <td><font size="2" color="#FFFFFF">���</font></td>

	</tr>
<%
do while not rs.eof and rowcount < pagesize
	rowcount = rowcount + 1
        if rs("flag")="F" then
            idx = "�O"
            sidx = rs("memno")
         else
            idx = "�_"
            sidx = rs("memno")
        end if 
        select case rs("code")

             
               case "E1"
                     cname ="�Ȧ楻��"
               case "E2"
                    cname ="�w�Х���"
               case "F1"
                    cname ="�Ȧ�Q��"  
               case "F2"
                    cname ="�w�ЧQ��"  
         
               case "A2"
		  cname ="�w�ЪѪ�"  
         end select 
                 
%>
  <tr bgcolor="#FFFFFF">
	
	<td><font size="2"><%=rs("memhkid")%></font></td>
	<td><font size="2"><%=rs("memno")%></font></td>
	<td><font size="2"><%=rs("memName")%></font></td>
        <td><font size="2"><%=cname%></font></td> 
	<td align="right"><font size="2"><%=formatNumber(rs("Bankin"),2)%></font></td>
	<td><a href="AutoAdkt.asp?del=<%=rs("memno")%>?<%=idx%>*<%=cname%>  " onclick="return confirm('����������H')"><font size="2"><%=idx%></font></a></td>

  </tr>
<%
	rs.movenext
loop
%>
</table>
</center>
</body>
</html>
