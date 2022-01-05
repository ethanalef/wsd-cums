<!-- #include file="../conn.asp" -->

<%
id  =request.form("memno")



chkdate = request.form("stdate")


set rs=conn.execute("select a.*,b.memcname,b.memGender from iva a,memmaster b where a.memno='"&id&"' and a.memno=b.memno  ")
if not rs.eof then
   lndate = right("0"&day(rs("lndate")),2)&"/"&right("0"&month(rs("lndate")),2)&"/"&year(rs("lndate"))
   shdate  = right("0"&day(rs("shdate")),2)&"/"&right("0"&month(rs("shdate")),2)&"/"&year(rs("shdate"))
   subttl = rs("slnamt") + rs("amount")
   total  = subttl - rs("shamt")
else
   response.redirect("noprint.asp")
end if
if request.form("output")="word" then
	Response.ContentType = "application/msword"
elseif request.form("output")="excel" then
	Response.ContentType = "application/vnd.ms-excel"

end if
  
%>
<html>
<head>
<title></title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
<center>
<br>
<%
   for i = 1 to 14
%>
<BR>
<%
   next
%>

<table border="0" cellpadding="0" cellspacing="0">
<tr>
     <td>&nbsp;</td>
     <td>&nbsp;</td>
</tr>
<tr>
     <td><b>Reconciliation Statement:</b></td>
     <td>&nbsp;</td>
</tr>
<tr>
     <td>&nbsp;</td>
     <td>&nbsp;</td>
</tr>
<tr>
     <td>Loan as at　<%=lndate%>　:　</td>
     <td align="right">$<%=formatnumber(rs("slnamt"),2)%></td>
</tr>

<tr>
     <td>Add: <%=rs("nmonth")%> months & <%=rs("nday")%> days interest</td>
     <td align="right"><u>$<%=formatnumber(rs("amount"),2)%></y></td>
</tr>
<tr>
     <td>&nbsp;</td>
      <td align="right">$<%=formatnumber(subttl,2)%></td>
</tr>
<tr>
     <td>Less: Share as at　<%=shdate%>　:　</td>
     <td align="right"><u>$<%=formatnumber(rs("shamt"),2)%></u></td>
</tr>
<tr>
     <td>&nbsp;</td>
      <td align="right">$<%=formatnumber(total,2)%></td>
</tr>
<tr>
     <td>&nbsp;</td>
      <td align="right">==========</td>
</tr>
</table>
<%
set rs=nothing
conn.close
set conn=nothing
%>
