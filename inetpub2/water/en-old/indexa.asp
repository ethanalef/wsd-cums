<!-- #include file="conn.asp" -->


<%

if request("From") = Request.ServerVariables("script_name") and request.form("username") <> "" then
	set reg = new regexp
	reg.pattern="[^a-zA-Z0-9]"
	reg.Global = True
	username=reg.replace(request("username"),"")
	password=reg.replace(request("password"),"")
	
	If ValidateUser(username, password) Then
		
		If LoadData(session("username")) Then
			Response.redirect "en/main.asp"
		else
			msg = "Error Occured While Loading User Rights"
		End If
		
	Else
		msg = "Login Failed"
	End If
	
		

    
end if

%>
<html>
<head>
<title>水務署員工儲蓄互助社系統</title>
<meta http-equiv="content-type" content="text/html; charset=big5">
<meta name="google-site-verification" content="+nxGUDJ4QpAZ5l9Bsjdi102tLVC21AIh5d1Nl23908vVuFHs34="/>
<link href="main.css" rel="stylesheet" type="text/css">
<script language=JavaScript>
<!--
function validating()
{
    if(document.login.username.value=="" || document.login.password.value=="")
    {
        alert("Please fill in both Username and Password");
        return false;
    }else{
        return true;
    }
}
//-->
</script>
</head>
<body bgcolor="#ffffff" leftMargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="login.username.focus()">
<div align="center">
<center>
<form method="POST" name="login" action="<% =Request.servervariables("script_name") %>" onSubmit="return validating()">
<input type=hidden name="From" value="<% =Request.servervariables("script_name") %>">
<br><br>
<font face="arial, helvetica, sans-serif" size="5" color="#336699"><b>水務署員工儲蓄互助社系統</b></font><br>
<font face="arial, helvetica, sans-serif" size="4" color="#336699"><b>Water Supplies Department Staff Credit Union<br>Membership, Accounting, Savings and Loans Software</b></font>
<br><br>
<img src="images/image002.gif" broder="0">
<br><br>
<font face="arial, helvetica, sans-serif" size="3" color="#000000">請輸入名稱及密碼登入系統<br>Please Login With Your Username and Password</font>
<br><br>
<table border="0" cellpadding="0" cellspacing="0">
    <tr>
        <td width="130"><b>名稱 Username</b></td>
        <td width="170"><input type="text" name="username" size="20"></td>
        <td width="100">&nbsp;</td>
    </tr>
    <tr>
        <td><b>密碼 Password</b></td>
        <td><input type="password" name="password" size="20"></td>
        <td><input type="submit" value="登入 Login"></td>
    </tr>
    <tr>
        <td colspan=2 height=60>
<%  if msg <> "" then %>
            <center><font color="#0000ff"><b><%= msg%></b></font></center>
<% end if %>
        </td>
    </tr>
</table>
<br><br>
<font size="2">Best Viewed With Microsoft Internet Explorer 5.0 or Higher</font>
<font size="2">Credur Union Ver 1.01.200604</font>
</form>
</center>
</div>
</body>
</html>



<%
' Function to validate user
Function ValidateUser(Username, Password)
	ValidateUser = False
	Dim CaseSensitive, AdminUsername, AdminPassword, SQL
	CaseSensitive = False ' Modify case sensitivity here
	If Not ValidateUser Then
		SQL = "select * from loginUser where username ='" & username & "'"
    		Set rs = Server.CreateObject("ADODB.Recordset")
    		rs.open SQL, conn,1 ,3

			If Not rs.eof Then
				If CaseSensitive Then
					ValidateUser = (rs("password") = Password)
				Else
					ValidateUser = (LCase(rs("password")) = LCase(Password))
				End If
				If ValidateUser Then
					session.timeout = 1200
                                Response.Cookies("userLevel")=rs("userLevel")
                                Response.Cookies("username")=rs("username")
        			session("userLevel") = rs("userLevel")
        			session("username") = rs("username")
					session("UID") = rs("uid")
        			session("workstart")=now
				End If
			End If
			rs.Close
			Set rs = Nothing
	End If
End Function






'-------------------------------------------------------------------------------
' Function LoadData
' - Load Data based on Key Value
' - Variables setup: field variables

Function LoadData(username)
	Dim rs, sSql
	


	sSql = "Select * FROM userRights WHERE Username = '"& username &"' "

	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open sSql, conn
	If rs.Eof Then
		LoadData = False
	Else
		LoadData = True
		rs.MoveFirst

		' Get the field contents
		
		Session("userRight_x_Member1") = rs("Member1")
		Session("userRight_x_Member2") = rs("Member2")
		Session("userRight_x_Member3") = rs("Member3")
                Session("userRight_x_Member4") = rs("Member4")
		Session("userRight_x_Member5") = rs("Member5")
		Session("userRight_x_Member6") = rs("Member6")
		Session("userRight_x_Member7") = rs("Member7")
		Session("userRight_x_Loan1") = rs("Loan1")
		Session("userRight_x_Loan2") = rs("Loan2")
		Session("userRight_x_Loan3") = rs("Loan3")
		Session("userRight_x_Loan4") = rs("Loan4")
		Session("userRight_x_Loan5") = rs("Loan5")
		Session("userRight_x_Loan6") = rs("Loan6")
		Session("userRight_x_Loan7") = rs("Loan7")
		Session("userRight_x_Loan8") = rs("Loan8")
		Session("userRight_x_Loan9") = rs("Loan9")
		Session("userRight_x_Loan10") = rs("Loan10")
                Session("userRight_x_Loan11") = rs("Loan11")
                Session("userRight_x_Loan12") = rs("Loan12")
		Session("userRight_x_cLoan1") = rs("cLoan1")
		Session("userRight_x_cLoan2") = rs("cLoan2")
		Session("userRight_x_cLoan3") = rs("cLoan3")
                Session("userRight_x_cLoan4") = rs("cLoan4")
                Session("userRight_x_cLoan5") = rs("cLoan5")
               Session("userRight_x_cLoan6") = rs("cLoan6")
                Session("userRight_x_cLoan7") = rs("cLoan7")
                 Session("userRight_x_cLoan8") = rs("cLoan8")
		Session("userRight_x_AutoPay1") = rs("AutoPay1")
		Session("userRight_x_AutoPay2") = rs("AutoPay2")
		Session("userRight_x_AutoPay3") = rs("AutoPay3")
		Session("userRight_x_AutoPay4") = rs("AutoPay4")
		Session("userRight_x_AutoPay5") = rs("AutoPay5")
		Session("userRight_x_AutoPay6") = rs("AutoPay6")
		Session("userRight_x_AutoPay7") = rs("AutoPay7")
		Session("userRight_x_AutoPay8") = rs("AutoPay8")
                Session("userRight_x_AutoPay9") = rs("AutoPay9")
		Session("userRight_x_AutoPay10") = rs("AutoPay10")
		Session("userRight_x_AutoPay11") = rs("AutoPay11")
		Session("userRight_x_AutoPay12") = rs("AutoPay12")
                Session("userRight_x_AutoPay13") = rs("AutoPay13")
                              
                 Session("userRight_x_MemAcct1") = rs("MemAcct1")
		Session("userRight_x_Saving1") = rs("Saving1")
		Session("userRight_x_Saving2") = rs("Saving2")
		Session("userRight_x_Saving3") = rs("Saving3")
		Session("userRight_x_Saving4") = rs("Saving4")
		Session("userRight_x_Saving5") = rs("Saving5")
		Session("userRight_x_Saving6") = rs("Saving6")
		Session("userRight_x_Saving7") = rs("Saving7")
		Session("userRight_x_Saving8") = rs("Saving8")
		Session("userRight_x_Saving9") = rs("Saving9")
                Session("userRight_x_Saving10") = rs("Saving10")  
		Session("userRight_x_Saving11") = rs("Saving11")
                Session("userRight_x_Saving12") = rs("Saving12")             		
		Session("userRight_x_Reporting1") = rs("Reporting1")
		Session("userRight_x_Reporting2") = rs("Reporting2")
		Session("userRight_x_Reporting3") = rs("Reporting3")
		Session("userRight_x_Reporting4") = rs("Reporting4")
		Session("userRight_x_Reporting5") = rs("Reporting5")
		Session("userRight_x_Reporting6") = rs("Reporting6")
		Session("userRight_x_Reporting7") = rs("Reporting7")
		Session("userRight_x_Reporting8") = rs("Reporting8")
		Session("userRight_x_Reporting9") = rs("Reporting9")
		Session("userRight_x_Reporting10") = rs("Reporting10")
		Session("userRight_x_Reporting11") = rs("Reporting11")
                Session("userRight_x_Reporting12") = rs("Reporting12")
                Session("userRight_x_Reporting13") = rs("Reporting13")
                Session("userRight_x_Reporting14") = rs("Reporting14")
                Session("userRight_x_Reporting15") = rs("Reporting15")
                Session("userRight_x_Reporting16") = rs("Reporting16")
		Session("userRight_x_Reporting17") = rs("Reporting17")
                Session("userRight_x_Reporting18") = rs("Reporting18")
                Session("userRight_x_Reporting19") = rs("Reporting19")
                Session("userRight_x_Statist1") = rs("Statist1")
                Session("userRight_x_Statist2") = rs("Statist2")
                Session("userRight_x_Statist3") = rs("Statist3")  
		Session("userRight_x_Other4") = rs("Other4")
		Session("userRight_x_Other3") = rs("Other3")
		Session("userRight_x_Other2") = rs("Other2")
		Session("userRight_x_Other1") = rs("Other1")
		Session("userRight_x_Other4") = rs("Other4")
		Session("userRight_x_Other5") = rs("Other5")
		Session("userRight_x_Other6") = rs("Other6")
	End If
	rs.Close
	Set rs = Nothing
End Function
%>










