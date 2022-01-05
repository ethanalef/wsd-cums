<!-- #include file="../conn.asp" -->
<!-- #include file="../addUserLog.asp" -->
<%

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

arrLevel = Array("Inactive","Member","Operator","Supervisor","Administrator","Auditor","Preview")

if request.form("back") <> "" then
	response.redirect "menu.asp"
end if


     


if request.form("bye") <>""  then
        id=""
 	For Each Field in Request.Form
		TheString = Field & "=id "
		Execute(TheString)
	Next
       response.redirect "chgpass.asp"
end if
if request.form("Search") <>""  then
	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
   	Next
        msg = ""
        tempass = password    
	set reg = new regexp
	reg.pattern="[^a-zA-Z0-9]"
	reg.Global = True
	username=reg.replace(request("username"),"")
	password=reg.replace(request("password"),"")
	If Not ValidateUser(username, password) Then

           msg = "用戶名稱不存在"		
        else


                id =  username
                pass = "readonly"
	End If
else
if request.form("action") <> "" then

    
        
	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
   	Next
        msg = ""

	If ValidateUser(username, password) Then

           msg = "用戶名稱已經存在"		

	End If
        if msg="" then    	
 
		conn.begintrans
                if tempass<> spassworsd and spassword<>"" then
                   conn.execute("update loginuser set password='"&spassword&"' where username='"&username&"' ")
                end if
  
		conn.committrans
		msg = "紀錄已更新"
                id=""
                username=""
        end if
else
   id = ""
   pass = ""
end if
end if



%>
<html>
<head>
<title>用戶管理-更改密碼</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
<!--

function popup(filename){
  window.open (filename,'pop','width=500,height=550,statusbar=no,toolbar=no,resizable,scrollbars,dependent')
}

function formatNum(numform){
  if (isNaN(numform.value)||numform.value<0)
    return false;
  else
    return true;
}

function valDate(M, D, Y){
  Months= new Array(31,28,31,30,31,30,31,31,30,31,30,31);
  Leap  = false;
  if((Y % 4 == 0) && ((Y % 100 != 0) || (Y %400 == 0)))
    Leap = true;
  if((D < 1) || (D > 31) || (M < 1) || (M > 12) || (Y < 0))
    return false;
  if((D > Months[M-1]) && !((M == 2) && (D > 28)))
    return false;
  if(!(Leap) && (M == 2) && (D > 28))
    return false;
  if((Leap)  && (M == 2) && (D > 29))
    return false;
  return true;
};



function formatDate(dateform){

	    formObj=document.form1;    
            sMn = parseInt(formObj.lastmonth.value)
            sYr = parseInt(formObj.lastyear.value)
            spass   = parseInt(formObj.spass.value)
          

  cDate = dateform.value;
  dSize = cDate.length;
 
  var strValidChars = "0123456789/";
  var strChar = "";

 
   for (i = 0; i < dSize ; i++)
      {  
      strChar = cDate.substr(i,1);
      if ( strValidChars.indexOf(strChar) == -1)
         { 
         return false ;   
          }
      }

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

      Mn = strM
      Yr = strY
      if (((Mn<=sMn)&&(Yr=sYr))||(Yr<sYr)){
         return false ;
      }else{      
         return true;
      }

  }
}



function checkpass(){
         formObj=document.form1;
         pass1 = formObj.spassword.value
         pass2 = formObj.password1.value
         if ( pass1!=pass2){
         alert(" 重入新密碼不符!") 
         formObj.password1.value = ""
         formObj.password1.focus()
         }
           
}

function validating(){
	formObj=document.form1;
	reqField="";
	placeFocus=false;

	if (formObj.username.value==""){
		reqField=reqField+", 用戶名稱";
		if (!placeFocus)
			placeFocus=formObj.memNo;
	}


    if (reqField){
        if (reqField.lastIndexOf(",")==0)
	        reqField = "請填入"+reqField.substring(2);
        else
	        reqField = "請填入"+reqField.substring(2,reqField.lastIndexOf(","))+'及'+reqField.substring(reqField.lastIndexOf(",")+2);
        alert(reqField);
        placeFocus.focus();
        return false;
    }else{
        return true;
    }
}
//  -->
</script>
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="form1.username.focus()">
<!-- #include file="menu.asp" -->
<%if msg<>"" then %>
<div><center><font size="3"><%=msg%></font></center></div>
<% end if%>

<br>
<center>
<form name="form1" method="post" action="chgpass.asp">
<input type="hidden" name="tempass" value="<%=tempass%>">
<div><center><font size="3">用戶管理-更改密碼</font></center></div>
<br>
<table border="0" cellpadding="0" cellspacing="0">

    <tr>
        <td width="130" align="right">使用者名稱</td>
        <td width="10">&nbsp</td>
        <td width="170" ><input type="text" name="username" value ="<%=username%>" size="20"  <%=pass%> ></td>
        <td width="100" >&nbsp;</td>
    </tr>
<%
   if id = "" then
%>
    <tr>
         <td width="130" align="right"><b>密碼</b></td>
        <td width="10">&nbsp</td>
        <td><input type="password" name="password" size="20" maxlength="20"    >

         <input type="submit" value="搜尋" name="Search" class ="Sbttn"> 
         <input type="submit" value="返回" name="back" class="sbttn">
        </td>
        <td width="100">&nbsp;</td>
    </tr>
<%
   end if

  if id <>"" then
%>
    <tr>
        <td width="130" align="right">新密碼</td>
        <td width="10">&nbsp</td>
        <td><input type="password" name="spassword" size="50" maxlength="50"></td>
         <td width="100">&nbsp;</td>
    </tr>
    <tr>
	 <td width="150" align="right">重入新密碼</td>	
        <td width="10">&nbsp</td>
	<td><input type="password" name="password1" value="<%=password1%>" size="20" maxlength="20" onblur="{checkpass()}" ></td>
        <td width="100">&nbsp;</td>
   </tr>

	<tr>
		<td colspan="9" align="right" valign="middle">
			
			<input type="submit" value="儲存" onclick="return validating()&&confirm('確定儲存?')" name="action" class="sbttn">
			
<%if uid="" then %>
		        <input type="submit" value="取消" name="bye"  class="sbttn">
                        <input type="button" value="返回" name="back" class="sbttn">
<%end if %>
			
		</td>
	</tr>
<%
   end if
%>
</table>
</form>
</center>
</body>
</html>