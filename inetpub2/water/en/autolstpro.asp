<!-- #include file="../conn.asp" -->

<!-- #include file="navigator.asp" -->
<%
if request("del")<>"" then
         conn.begintrans
      id = request("del")
      if instr(id,"�_")> 0 then
         memno = left(id,instr(id,"�_")-1)
 
         conn.execute("update autopay set pflag = 1 where memno='"&memno&"' ")
      else
      if instr(id,"�O")> 0 then
        memno = left(id,instr(id,"�O")-1)
         conn.execute("update autopay set pflag = 0 where memno='"&memno&"' ")
      end if
      end if
     

      conn.committrans 
    
   
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


IF REQUEST("NPAGE") <> "" OR REQUEST("UPAGE") <>"" THEN
   SQL_FILTER  =SESSION("STRSQL")
END IF

sql = "select  a.*,b.memname,b.memcname from autopay a,memmaster b where a.memno=b.memno and a.flag='F'  " & sql_filter

set rs = server.createobject("ADODB.Recordset")
rs.open sql, conn, 1,1




if not rs.eof then
	if request("npage") <> "" then          
           pageno = session("cpageno")+1
           curpage = pageno
        
	else
		pageno = 1
                curpage = 0
	end if
	if request("upage") <> "" then          
           pageno = session("cpageno")-1
           curpage = pageno
        

	end if
	rs.pagesize = 10
	pagesize=rs.pagesize
	rs.absolutepage = pageno
	recordcount=rs.recordcount
	pagecount = rs.pagecount
        rowcount  = 0
        session("cpageno") = pageno
        SESSION("STRSQL")=sql_filter

end if
%>
<html>
<head>
<title>�Ȧ�۰���b���ĳq���ѦC�L</title>
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
''		<td bgcolor="#FFFFFF"><a href="autolstproDetail.asp"><font size="2">�s�W</font></a></td>
</script>
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0">
<!-- #include file="menu.asp" -->
<br>
<center>
<%if msg<>"" then response.write "<font color=red>"&msg&"</font>" end if%>
<form name="form1" method="post" action="autolstpro.asp">
<div><center><font size="3">�Ȧ�۰���b���ĳq���ѦC�L</font></center></div>

�����m�W : <input type="text" name="memNo" value="<%=memNo%>" size="6">
�����m�W : <input type="text" name="memName" value="<%=memName%>" size="10">

<input type="submit" name="memSearch" value="�j�M" class="sbttn" >
</form>

<table border="0" cellspacing="1" cellpadding="4" align="center" bgcolor="336699">
	<tr bgcolor="#330000" align="center">
		
		<td><font size="2" color="#FFFFFF">�����s��</font></td>
		<td><font size="2" color="#FFFFFF">�����m�W</font></td>
		<td><font size="2" color="#FFFFFF">�C�L</font></td>

	</tr>
<%
do while not rs.eof and rowcount < pagesize
	rowcount = rowcount + 1
        if rs("pflag") then
           idx ="�O"

        else
            idx ="�_"
        end if
%>
  <tr bgcolor="#FFFFFF">
	
	
	<td><font size="2"><%=rs("memNo")%></font></td>
	<td><font size="2"><%=rs("memName")%></font></td>
	<td align="right"><font size="2"><%=idx%></font></td>
<%if idx ="�O" then%>
	<td><a href="autolstpro.asp?del=<%=rs("memNo")%><%=idx%>" onclick="return confirm('�����C�L?')"><font size="2">�����C�L</font></a></td>
<%else%>
        <td><a href="autolstpro.asp?del=<%=rs("memNo")%><%=idx%>" onclick="return confirm('�T�w�C�L?')"><font size="2">�T�w�C�L</font></a></td>
<%end if%>
  </tr>
<%
	rs.movenext
loop
%>
</table>
<%if session("cpageno")>1 then%>
    <a href="autolstpro.asp?upage=upage<font size="2">�W�@��</font></a>
<%end if%>
<%if session("cpageno")< pagecount then%>
<a href="autolstpro.asp?npage=npage<font size="2">�U�@��</font></a>
<%end if %>

</body>
</html>
