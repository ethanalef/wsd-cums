<!-- #include file="../conn.asp" -->
<!-- #include file="navigator.asp" -->
<%

key = ""
stylefield ="" 


if request.form("Search")<> ""  then
	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next


select case  status
       case  "�����s��"
             stylefield ="where  memNo like '" & key & "%'  order by memno "
             STATUS = "1" 
       case  "�^��m�W"
             stylefield ="where  memname like '" & ucase(key) & "%' order by memno "
             STATUS ="2"
       case  "����m�W"
	     stylefield ="where  memcname like '" & key & "%' order by memno "
             STATUS = "3"
       case  "�����ҽs��"
	     stylefield ="where  memhkid like '" & key & "%' order by memno "
             STATUS = "4"
end select

END IF
response.write( request("page"))
if request("page")  = "" then
   Session("xstylefield")= stylefield
else
   stylefield = Session("xstylefield")
    Session("xstylefield")= stylefield
end if
response.write(Session("xstylefield"))
SQL = "select memNo,memname,memcname,memhkid  from memmaster   "& stylefield 
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn, 3
if rs.recordcount=1 then
%>
<script language="JavaScript">
<!--
	window.opener.document.form1.memNo.value='<% =rs("memNo") %>'
        window.opener.document.form1.id.value='<% =rs("memNo") %>'
 
	window.opener.document.form1.Search.click()        
	
	window.close();
//-->
</script>
<%
	response.end
end if

if not rs.eof then
	if request("page") <> "" then
		pageno = cint(request("page"))
	else
		pageno = 1
	end if
	rs.pagesize = 15
	pagesize=rs.pagesize
	rs.absolutepage = pageno
	recordcount=rs.recordcount
	pagecount = rs.pagecount
end if


%>
<html>
<head>
<meta1 HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<title>Member Popup</title>
<link href="../main.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
<!--
function PostCode(memNo,memName,memcName ) 
{
 
	window.opener.document.form1.memNo.value=memNo;
        window.opener.document.form1.id.value=memNo;

	window.opener.document.form1.Search.click() 	
	window.close();
}
//-->
</script>

</head>
<body leftMargin="0" topMargin="0" marginheight="0" marginwidth="0"  >
<div align="center">
<center>
<table border="0" cellpadding="0" cellspacing="0" width="100%" bgcolor="#87CEEB">
	<form  method="POST" action="<%=request.servervariables("script_name")%>" name="form">
    <tr>
        <td height="25">
			&nbsp;&nbsp;<b>�j�M</b>
			<input type="text" name="key"  size="20" maxlength="50">		
			<select name="status">			
			<option<%if status="1" then response.write " selected" end if%>>�����s��    </option>
                        <option<%if status="2" then response.write " selected" end if%>>�^��m�W</option>
                        <option<%if status="3" then response.write " selected" end if%>>����m�W</option>
                        <option<%if status="4" then response.write " selected" end if%>>�����ҽs��</option>

			</select>		
			<input type="submit" name="Search" value="�T�w" class="sbttn" >
        </td>
        <td align="right">
            <font face="arial" style="color: black; font: bold" size="3">�j�M����</font>&nbsp;&nbsp;
        </td>
    </tr>
	</form>
</table>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<%
if key <> "" then
%>
	<tr>
		<td align="center">"<% =key %>" ���j�M���G</td>
	</tr>
<%
end if
%>
	<tr>
		<td>
<%
if not rs.eof then
	
%>
		<table border="0" cellpadding="3" cellspacing="1" width="100%" bgcolor="#336699">
<%
	do while not rs.eof and rowcount < rs.pagesize
		rowcount = rowcount + 1
%>
			<tr>
				<td bgcolor="#ffffff"><a href="JavaScript:PostCode('<%=rs(0)%>','<%=rs(1)%>','<%=rs(2)%>','<%=rs(3)%>')"><%=rs(0)%></a></td>
				<td bgcolor="#ffffff"><% =rs(1) %></a></td>
				<td bgcolor="#ffffff"><% =rs(2) %></a></td>
				<td bgcolor="#ffffff"><% =rs(3) %></a></td>

				</tr>
<%
		rs.movenext
	loop
%>
		</table>
<%
	navigator(request.servervariables("script_name")&"?key="&key&status)
	response.write "<p>"
else
%>
		<br><p align="center"><font size="4">�L�ŦX����</font></p>
<%
end if
%>
		</td>
	</tr>
</table>
</center>
</div>
</body>
</html>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>