<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<%
For Each Field in Request.Form
	TheString = "[" & Field & "] = Request.Form(""" & Field & """)"
	Execute(TheString)
Next

f = ""
a = split(request("TS"),",",-1,1)
if ubound(a) < 0 then
	response.redirect "acBal.asp"
end if

for i = 0 to ubound(a)
	f = f & "'"&trim(A(i))&"',"
next
f = left(f,len(f)-1)

SQL = "select * from glControl"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn
acYear=rs("acYear")
rs.close

if request("mYear")="this" then
	mShr = "thisShrBal"
	mLoan = "thisLoanBal"
else
	acYear=acYear-1
	mShr = "lastShrBal"
	mLoan = "lastLoanBal"
end if

mStart=int(mStart)
mEnd=int(mEnd)

if mStart<=4 then
	m=mStart+8
	y=acYear
else
	m=mStart-4
	y=acYear+1
end if
startDate=cdate(y&"/"&m&"/1")
if mEnd<=4 then
	m=mEnd+8
	y=acYear
else
	m=mEnd-4
	y=acYear+1
end if
if m=12 then m=1:y=y+1 else m=m+1 end if
endDate=cdate(y&"/"&m&"/1")-1

SQl = "select * from memMaster where deleted=0 and memNo between "&mFrom&" and "&mTo&" and memSection In ("&f&") order by memSection,memNo"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn
%>
<html>
<head>
<title>Print Balance Statement</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body leftMargin="0" topmargin="0" marginheight="0" marginwidth="0">
<pre>
<%
cycle=1
do while not rs.eof
	response.write "           "&right("0"&day(startDate),2)&"/"&right("0"&month(startDate),2)&"/"&right(year(startDate),2)&"      "&right("0"&day(endDate),2)&"/"&right("0"&month(endDate),2)&"/"&right(year(endDate),2)&space(55)&rs("memSection")&space(6)&request.form("mDate")&vbCr&vbCr&vbCr
	response.write "        "&left(rs("memName")&"                              ",30)&"               "&rs("memNo")
	response.write "      Dividend : "&formatNumber(rs("dividend"),2)&vbCr&vbCr&vbCr&"       "
	for idx=1 to 6
		if idx>=mStart and idx<=mEnd then
			response.write right("                 "&formatNumber(rs(mShr&idx),2),17)
		else
			response.write "             0.00"
		end if
	next
	response.write vbCr&vbCr&vbCr&"       "
	for idx=7 to 12
		if idx>=mStart and idx<=mEnd then
			response.write right("                 "&formatNumber(rs(mShr&idx),2),17)
		else
			response.write "             0.00"
		end if
	next
	response.write vbCr&vbCr&vbCr&"       "
	for idx=1 to 6
		if idx>=mStart and idx<=mEnd then
			response.write right("                 "&formatNumber(rs(mLoan&idx),2),17)
		else
			response.write "             0.00"
		end if
	next
	response.write vbCr&vbCr&vbCr&"       "
	for idx=7 to 12
		if idx>=mStart and idx<=mEnd then
			response.write right("                 "&formatNumber(rs(mLoan&idx),2),17)
		else
			response.write "             0.00"
		end if
	next
	if cycle=4 then
		response.write vbCr
		cycle=1
	else
		cycle=cycle+1
		response.write vbCr&vbCr&vbCr
	end if
	rs.movenext
loop
%>
</pre>
</body>
</html>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
