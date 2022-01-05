<!-- #include file="../conn.asp" -->

<%
   server.scripttimeout = 1800
   yyy = request.form("myear")

   yyy = "2006"

   stdate  = yyy&".1.1"
   eddate = yyy&"/12/31"
   xyy = cint(yyy)   



dim   age(6)
dim   share(9)
dim   loan(9) 

aettl = 0
shtl = 0
lnttl = 0





SQl = "select memno,membday  from memmaster  where  ( memdate<> null or  memdate <='"&eddate&"' ) and (wdate is null or wdate>'"&eddate&"' )  "
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn, 1

   do while not rs.eof 


       xage = xyy - year(rs("membday"))
       mm = (12 -  month(rs("membday")) )
       if mm >=6 then
          xage = xage + 1
       end if

       agettl = agettl + 1 
       if xage>=18 and xage <= 30 then
          age(1)= age(1)+1
       end if
       if xage>=31 and xage <=40 then   
         age(2)= age(2) + 1
       end if
       if xage>=41 and xage <= 50 then
          age(3)= age(3)+1
       end if
       if xage>=51 and xage <=60 then   
         age(4)= age(4) + 1
       end if
       if xage>=61 and xage <= 70 then
          age(5)= age(5)+1
       end if
       if xage>=71  then   
         age(6)= age(6) + 1
       end if            


   rs.movenext
   loop
   rs.close


for i = 1 to 6
    response.write(age(i))
    response.write("**")
next

response.end

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
	objFile.Write "                           水務署員工儲蓄互助社"
	objFile.WriteLine ""
	objFile.Write "                              年結分析報告 －"&yyy
	objFile.WriteLine ""




	objFile.Close
	
	set rs=nothing
	conn.close
	set conn=nothing
	response.redirect "../txt/"&session("username")&".txt"
end if
%>
<html>
<head>
<title>Delinquent Report</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<LINK href="../main.css" rel=STYLESHEET type=text/css>
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
<center>
<table border="0" cellpadding="0" cellspacing="0">
	<tr height="30" valign="top" align="center">
		<td colspan="15"><font size="5">水務署員工儲蓄互助社</font></td>
	</tr>
	<tr height="35" valign="top" align="center">
		<td colspan="15"><font size="5">呆帳列表</font></td>
	</tr>
	<tr height="35" valign="top" align="center">
		<td colspan="15"><font size="5">已過日期 <%=nday%> 日</font></td>
	</tr>
	<tr height="15" valign="bottom">
		<td width="80"><b>社員名稱</b></td>
		<td width="200"><b>社員名稱</b></td>
		<td width="130" align="right"><b>貸款編號</b></td>
		<td width="130" align="right"><b>貸款總額</b></td>
		<td width="130" align="right"><b>貸款結欠</b></td>
	</tr>
	<tr><td colspan=5><hr></td></tr>

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
