<!-- #include file="../conn.asp" -->
<% ' <!-- #include file="../CheckUserStatus.asp" --> %>
<%

server.scripttimeout = 1800

ph1 = "PH"
ph2 = "F"
ph3 = "F01"
ph4 = space(12)
ph5 = year(date()) & right("0" & month(date()), 2) & right("0" & day(date()), 2)
ph6 = left("024062010001" & space(35), 35)
ph7 = "HKD"
'ph8
set rs = conn.execute("select count (*) from (select memno from dividend where bank = 'B' AND dividend > 0 group by MEMNO) as mem")
if not rs.eof then
	ph8 = rs(0)
else
	ph8 = 0
end if
rs.Close
ph8 = right("000000" & ph8, 7)
'ph9
set rs = conn.execute("select round(sum(dividend), 2) from Dividend where bank = 'B'")
if not rs.eof then
 ph9 = rs(0) * 100
end if
rs.close
ph9 = right("00000000000000000" & ph9, 17)
ph10 = space(1)
ph11 = space(311)

header = ph1 & ph2 & ph3 & ph4 & ph5 & ph6 & ph7 & ph8 & ph9 & ph10 & ph11 & vbCrLf

detail = ""
pd1 = "PD"
pd3 = "BBAN"
pd7 = space(35)
pd9 = space(130)
set rs = server.createobject("ADODB.Recordset")
sql  = "select a.memno, a.dividend, b.memcname, b.memname, b.bnk, b.bch, b.bacct from dividend a, memmaster b where a.memno = b.memno and a.bank='B' order by a.memno, b.memcname, b.memname, b.bnk, b.bch, b.bacct"
rs.open sql, conn
do while  NOT rs.eof
	pd2 = rs("bnk")
	pd4 = left(rs("bch") & rs("bacct") & space(34), 34)
	pd5 = right("00000000000000000" & rs("dividend") * 100, 17)
	pd6 = left("NO" & right(space(5) & rs("memno"), 5) & space(34), 35)
	pd8 = left(UCase(rs("memname")) & space(140), 140)
	detail = detail & pd1 & pd2 & pd3 & pd4 & pd5 & pd6 & pd7 & pd8 & pd9 & vbCrLf
	rs.movenext
loop
set rs = nothing
conn.close
set conn = nothing


set fs = Server.CreateObject("Scripting.FileSystemObject")
set f = fs.CreateTextFile("c:\public\shpay.apc", true)
f.WriteLine header & detail
f.close
set f = nothing
set fs = nothing

RESPONSE.REDIRECT "COMPLETED.ASP"
%>

<html>
	<head>
		<title></title>
		<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
		<link href="../main.css" rel="stylesheet" type="text/css">
	</head>
	<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
	</body>
</html>