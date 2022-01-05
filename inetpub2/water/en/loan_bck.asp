<!-- #include file="../conn.asp" -->

<!-- #include file="navigator.asp" -->
<%
   username = session("username")
   userlevel = session("userlevel")
if request("del")<>"" then
	uid = request("del")
'	set rs = server.createobject("ADODB.Recordset")
	conn.execute("delete meetingnotes0 where appId="&uId)
	conn.execute("delete meetingnotes1 where appId="&uId)
	conn.execute("update loanApp set deleted=-1 where uId="&uId)
	msg = uId&" 已刪除"
'	rs.close
end if
xmon = month(date()) -4
xyr = year(date())
if xmon <= 0  then
   xmon = xmon + 12
   xyr  = xyr - 1
end if
mcdate = "01/"&right("0"&xmon,2)&"/"&xyr

For Each Field in Request.Form
	TheString = Field & "= Request.Form(""" & Field & """)"
	Execute(TheString)
Next
For Each Field in Request.querystring
	TheString = Field & "= Request.querystring(""" & Field & """)"
	Execute(TheString)
Next
if uid <> "" then
	sql_filter = sql_filter & " uid = '"&uid&"'  "
end if
if memNo <> ""  then
   if  sql_filter <>""  then            
	sql_filter = sql_filter & " and memNo = '"&memNo&"' "
   else
        sql_filter = sql_filter & " memNo = '"&memNo&"' "
   end if
end if
if memName <> "" then
   if  sql_filter <>""  then          
	sql_filter = sql_filter & " and memName like '%"&memName&"%' "
   else
        sql_filter = sql_filter & " memName like '%"&memName&"%'  "
   end if
end if
if date1 <> "" then
      if  sql_filter <>""  then  
	  sql_filter = sql_filter & " and  appDate >= '"&date1&"' and  "
      else
          sql_filter = sql_filter & " appDate >= '"&date1&"' and  "
      end if
end if
if date2 <> "" then
     if  sql_filter <>""  then  
	sql_filter = sql_filter & " and  appDate <= '"&date2&"'  "
     else
	sql_filter = sql_filter & " appDate <= '"&date2&"'  "
     end if 
end if


if approval <> "" then
        select case Approval
               case "批准申請"
                    
                    xapproval="Approved"
                    if  sql_filter <>""  then  
                         sql_filter = sql_filter & " and (firstapproval='"&xapproval&"' or SecondApproval ='"&xapproval&"' ) and deleted = '0' "    
                    else
                        sql_filter = sql_filter & " (firstapproval='"&xapproval&"' or SecondApproval='"&xapproval&"' ) and deleted = '0' "     
                    end if
               case "拒絕申請"
                    
	   	    xApproval="Rejected"
                       if  sql_filter <>""  then   
                          sql_filter = sql_filter & " and (firstapproval='"&xapproval&"' or SecondApproval='"&xapproval&"' ) and deleted = '0' "  
                       else
                          sql_filter = sql_filter & " (firstapproval='"&xapproval&"' or SecondApproval='"&xapproval&"' ) and deleted = '0' "   
                       end if 
               case "審理中"
                 
		    xApproval="Pending"	
                       if  sql_filter <>""  then   
                          sql_filter = sql_filter & " and (firstapproval='"&xapproval&"' or SecondApproval='"&xapproval&"' ) and deleted = '0' "  
                       else
                          sql_filter = sql_filter & " (firstapproval='"&xapproval&"' or SecondApproval='"&xapproval&"' ) and deleted = '0' "   
                       end if 
               case "取消申請"
                        
		  xApproval="cancel"	
                  if  sql_filter <>""  then   
                          sql_filter = sql_filter & " and (firstapproval='"&xapproval&"' or SecondApproval='"&xapproval&"' ) and deleted = '0' "  
                  else
                          sql_filter = sql_filter & " (firstapproval='"&xapproval&"' or SecondApproval='"&xapproval&"' ) and deleted = '0' "   
                  end if 
        end select 
else
        	       xApproval="Pending"	
                       if  sql_filter <>""  then   
                          sql_filter = sql_filter & " and (firstapproval='"&xapproval&"' or SecondApproval='"&xapproval&"' ) and deleted = '0' "  
                       else
                          sql_filter = sql_filter & " (firstapproval='"&xapproval&"' or SecondApproval='"&xapproval&"' ) and deleted = '0' "   
                       end if           
end if
sql = "select *,convert(char(10),appdate,102) as chkdate from loanApp where  " & sql_filter & " order by uid desc"
response.write(sql)
response.end

set rs = server.createobject("ADODB.Recordset")
rs.open sql, conn, 3

'if rs.recordcount=0 then
'	response.redirect "loanDetail.asp"
'end if

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
<title>貸款申請</title>
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
<input type="hidden" name="Approval" value="<%=Approval%>">
<br>
<center>
<%if msg<>"" then response.write "<font color=red>"&msg&"</font>" end if%>
<form name="form1" method="post" action="loan.asp">
申請序號 : <input type="text" name="uid" value="<%=uid%>" size="10">
日期 : <input type="text" name="date1" value="<%=date1%>" size="10" onblur="if(!formatDate(this)){this.value=''};">
- <input type="text" name="date2" value="<%=date2%>" size="10" onblur="if(!formatDate(this)){this.value=''};">
社員編號 : <input type="text" name="memNo" value="<%=memNo%>" size="6">
社員姓名 : <input type="text" name="memName" value="<%=memName%>" size="10">
批核申請 : <td width="100">
			<select name="Approval">
			<option>
			<option>批准申請
			<option>拒絕申請
			<option>審理中
                        <option>取消申請
			</select>
		
<input type="submit" value="搜尋" onclick="return validating()" class="sbttn">
</form>
<%if recordcount>pagesize then navigator("loan.asp?uid="&uid&"&date1="&date1&"&date2="&date2&"&memNo="&memNo&"&memName="&memName&"&Approval="&Approval) end if%>
<table border="0" cellspacing="1" cellpadding="4" align="center" bgcolor="336699">
	<tr bgcolor="#330000" align="center">
		<td><font size="2" color="#FFFFFF">申請序號</font></td>
		<td><font size="2" color="#FFFFFF">日期</font></td>
		<td><font size="2" color="#FFFFFF">社員編號</font></td>
		<td><font size="2" color="#FFFFFF">社員姓名</font></td>
		<td><font size="2" color="#FFFFFF">貸款額</font></td>
		<td><font size="2" color="#FFFFFF">貸款情況</font></td>
<%if session("userLevel")<>5 then%>
		<td bgcolor="#FFFFFF"><a href="loanDetail.asp"><font size="2">新增</font></a></td>
<%end if%>
	</tr>
<%
do while not rs.eof and rowcount < pagesize
if rs("chkdate") > mcdate then 
        appstatus = "等待批核"
	rowcount = rowcount + 1
        if rs("firstApproval") <> "" then
           select case rs("firstApproval")
                  case "Approved"
                       appstatus ="委員巳批准申請"
                  case "Rejected"
                       appstatus ="委員巳否決申請"
                  case "Pending"
                       appstatus ="委員在審理中"   
                  Case "cancel"
                       appstatus ="取消申請"   
            end select
       else
       if rs("SecondApproval") <>"" then
          select case rs("secondApproval") 
                 case "Approved"
                       appstatus ="董事巳批准申請"
                 case "Rejected"
                      appstatus ="董事巳否決申請"
                case "Pending"
                     appstatus ="董事在審理中"
                end select
       end if
       end if
       xappdate=right("0"&day(rs("appdate")),2)&"/"&right("0"&month(rs("appdate")),2)&"/"&year(rs("appdate"))
%>
  <tr bgcolor="#FFFFFF">
	<td><a href="loanDetail.asp?uid=<%=rs("uid")%>"><font size="2"><%=rs("uid")%></font></a></td>
	<td><font size="2"><%=xappdate%></font></td>
	<td><font size="2"><%=rs("memNo")%></font></td>
	<td><font size="2"><%=rs("memName")%></font></td>
	<td align="right"><font size="2"><%=formatNumber(rs("loanAmt"),2)%></font></td>
        <td><font size="2"><%=appstatus%></font></td>

  </tr>
<%
end if
	rs.movenext
loop
%>
</table>
</center>
</body>
</html>
