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

               case "�Ȧ�"
                   conn.execute("update dividend  set Bank='C'  where memno='"&smemno&"'  " )
               case "�䲼"
                   conn.execute("update dividend set Bank='B'  where memno='"&smemno&"'  ") 

  end select
end if

if request("sdel")<>"" then
 
  uid = request("Sdel")

  pos = instr(uid,"?")
  pos1 = instr(uid,"*")
  dif  = pos1 - pos-1
  smemno = left(uid,pos-1) 
  status = mid(uid,pos+1,dif)
  cname  = mid(uid,pos1+1,10)

  select case cname

               case "�O"
                   conn.execute("update dividend  set deleted=0   where memno='"&smemno&"'  " )
               case "�_"
                   conn.execute("update dividend set deleted =1  where memno='"&smemno&"'  ") 

  end select

  memno = ""
end if

pdamt = cint(request.form("minpaid"))
if pdamt > 0 then
conn.execute("update dividend set  bank = 'B' where dividend > "&pdamt&" and bank is null ")
conn.execute("update dividend set  bank = 'S' where dividend <= "&pdamt&" ")
conn.execute("update dividend set  DELETED = 0 where deleted <> '1' ")
pdamt = 0
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
    case "B" 
    sql_filter = sql_filter & " and a.bank <> 'B' "
   case "N"
    sql_filter = sql_filter & " and a.bank = 'N' "
end select



IF REQUEST("NPAGE") <> "" OR REQUEST("UPAGE") <>"" THEN
   SQL_FILTER  =SESSION("STRSQL")
END IF

 
sql = "select a.*,b.memcname,b.memhkid from Dividend a,memmaster b where  a.memno=b.memno  "& sql_filter 
set rs = server.createobject("ADODB.Recordset")

rs.open sql, conn, 3


if not rs.eof then
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
<title>�Ѯ��᤼���t�إ�</title>
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
<form name="form1" method="post" action="SeparatProc.asp">

�����Ҹ��X : <input type="text" name="memhkid" value="<%=memhkid%>" size="10">
�����s�� : <input type="text" name="memNo" value="<%=memNo%>" size="6">
�����m�W : <input type="text" name="memName" value="<%=memName%>" size="10">
���p(Y/N/A) : <input type="text" name="sflag" value="<%=sflag%>" size="1">
<input type="submit" value="�j�M"  onclick="return validating()" class="sbttn">
</form>
<%if recordcount>pagesize then navigator("SeparatProc.asp?memhkid="&memhkid&"&memNo="&memNo&"&memName="&memName&"&sflag="&sflag) end if%>
<table border="0" cellspacing="1" cellpadding="4" align="center" bgcolor="336699">
	<tr bgcolor="#330000" align="center">
	
		<td><font size="2" color="#FFFFFF">�����Ҹ��X</font></td>
		<td><font size="2" color="#FFFFFF">�����s��</font></td>
		<td><font size="2" color="#FFFFFF">�����m�W</font></td>
                <td><font size="2" color="#FFFFFF">���t���O</font></td>
                <td><font size="2" color="#FFFFFF">������</font></td>
		<td><font size="2" color="#FFFFFF">��b���B</font></td>
                <td><font size="2" color="#FFFFFF">���p</font></td>
                 <td><font size="2" color="#FFFFFF">�R��</font></td>
 
	</tr>
<%
do while not rs.eof and rowcount < pagesize
	rowcount = rowcount + 1
        select case  rs("Bank")
               case "B" 
                     idx = "�Ȧ�"
                     cname = idx
                     sidx = rs("memno")
              case "C"
                   idx = "�䲼"
                   cname = idx
                   sidx = rs("memno")
        end select
     
       IF  rs("deleted")= TRUE  THEN
                     xidx = "�O"
                     xname = xidx
                     sidx = rs("memno")
        ELSE
                   xidx = "�_"
                   xname = xidx
                   sidx = rs("memno")
        end IF
           
%>
  <tr bgcolor="#FFFFFF">
	
	<td><font size="2"><%=rs("memhkid")%></font></td>
	<td><font size="2"><%=rs("memno")%></font></td>
	<td><font size="2"><%=rs("memcName")%></font></td>
        <td><font size="2"><%=cname%></font></td> 
                <td><font size="2"><%=xname%></font></td> 
	<td align="right"><font size="2"><%=formatNumber(rs("Dividend"),2)%></font></td>
	<td><a href="SeparatProc.asp?del=<%=rs("memno")%>?<%=idx%>*<%=cname%>  " onclick="return confirm('��惡�����H')"><font size="2"><%=idx%></font></a></td>
        <td><a href="SeparatProc.asp?sdel=<%=rs("memno")%>?<%=xidx%>*<%=xname%>  " onclick="return confirm('�������R���H')"><font size="2"><%=xidx%></font></a></td>
  </tr>
<%
	rs.movenext
        loop
%>
</table>
</center>
</body>
</html>
