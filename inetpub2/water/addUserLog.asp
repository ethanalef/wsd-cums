<%
Sub addUserLog(action)
        set rsadd = server.createobject("ADODB.Recordset")
	sql = "select count(*) from userLog"
        rsadd.open sql, conn, 1, 1   	
	if rsadd(0) =0 then
		conn.execute("insert into userLog (uid,username,userLevel,actionDes,actionTime) values ( 1,'"&session("username")&"',"&session("userLevel")&",'"&action&"',getdate() )")
	else
              
		conn.execute("insert into userLog (uid,username,userLevel,actionDes,actionTime) select max(uid)+1,'"&session("username")&"',"&session("userLevel")&",'"&action&"',getdate() from userLog")
	end if
        rsadd.close
End Sub
%>