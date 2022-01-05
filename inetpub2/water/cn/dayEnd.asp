<%requiredLevel=2%>
<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<%
if request("process")<>"" then
	response.redirect "completed.asp"
end if
%>
<html>
<head>
<title>Day End</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" alink="#003399" link="#003399" vlink="#003399" bgcolor="#eeeef0">
<!-- #include file="menu.asp" -->
<br>
<center>
<h3>Day End</h3>
<form name="form1" method="post" action="<% =Request.servervariables("script_name") %>">
<input type="submit" name="process" value="Process">
</form>
</center>
</body>
</html>
