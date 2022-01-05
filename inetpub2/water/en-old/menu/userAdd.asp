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
if request.form("new") <> "" then
	response.redirect "UserAdd.asp"
end if


     


if request.form("bye") <>""  then
        id=""
 	For Each Field in Request.Form
		TheString = Field & "=id "
		Execute(TheString)
	Next
       response.redirect "menu.asp"
end if

if request.form("action") <> "" then

    
        
	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
   	Next
        msg = ""
	set reg = new regexp
	reg.pattern="[^a-zA-Z0-9]"
	reg.Global = True
	username=reg.replace(request("username"),"")
	password=reg.replace(request("password"),"")
	If ValidateUser(username, password) Then

           msg = "用戶名稱已經存在"		

	End If
        if msg="" then    		
		conn.begintrans
	        set rs = server.createobject("ADODB.Recordset")
		sql = "select top 1 * from loginUser order by uid desc"
		rs.open sql, conn, 2, 2		
		if not rs.eof then
		   id = rs("uid") + 1
		end if
		rs.addnew
		rs("username") = username
		rs("uid") = id
		addUserLog "Add User"
		rs("userLevel") = cdbl(userLevel)
		if password<>"" then rs("password") = password end if
		rs.update
                rs.close
		sql = "select top 1 * from UserRights order by username desc"
		rs.open sql, conn, 2, 2	
                if not rs.eof then	
		rs.addnew
		rs("username") = username
		rs("user_fk") = id
		rs("Member1") = Tsc11
		rs("Member2") = Tsc12
		rs("Member3") = Tsc13
		rs("Member4") = Tsc14
		rs("Member5") = Tsc15
		rs("Member6") = Tsc16
		rs("Member7") = Tsc17
                rs("Loan1") = Tsc21
                rs("Loan2") = Tsc22   
                rs("Loan3") = Tsc23
                rs("Loan4") = Tsc24
                rs("Loan5") = Tsc25
                rs("Loan6") = Tsc26   
                rs("Loan7") = Tsc27
                rs("Loan8") = Tsc28
                rs("Loan9") = Tsc29
                rs("Loan10") = Tsc2A   
                rs("Loan11") = Tsc2B
                rs("Loan12") = Tsc2C
                rs("cLoan1") = Tsc31
                rs("cLoan2") = Tsc32   
                rs("cLoan3") = Tsc33
                rs("cLoan4") = Tsc34
                rs("cLoan5") = Tsc35
                rs("cLoan6") = Tsc36   
                rs("cLoan7") = Tsc37 
                rs("cLoan8") = Tsc38 
                rs("Autopay1") = Tsc41
                rs("Autopay2") = Tsc42
                rs("Autopay3") = Tsc43
                rs("Autopay4") = Tsc44
                rs("Autopay5") = Tsc45
                rs("Autopay6") = Tsc46
                rs("Autopay7") = Tsc47
                rs("Autopay8") = Tsc48
                rs("Autopay9") = Tsc49
                rs("Autopay10") = Tsc4A
                rs("Autopay11") = Tsc4B
                rs("Autopay12") = Tsc4C
                rs("Autopay13") = Tsc4D
                rs("Saving1") = Tsc51
                rs("Saving2") = Tsc52
                rs("Saving3") = Tsc53
                rs("Saving4") = Tsc54
                rs("Saving5") = Tsc55
                rs("Saving6") = Tsc56
                rs("Saving7") = Tsc57
                rs("Saving8") = Tsc58
                rs("Saving9") = Tsc59
                rs("Saving10") = Tsc5A
                rs("Saving11") = Tsc5B
                rs("Saving12") = Tsc5C
                rs("MemAcct1") =Tsc61
                rs("Reporting1") = Tsc71
                rs("Reporting2") = Tsc72
                rs("Reporting3") = Tsc73
                rs("Reporting4") = Tsc74
                rs("Reporting5") = Tsc75
                rs("Reporting6") = Tsc76
                rs("Reporting7") = Tsc77
                rs("Reporting8") = Tsc78
                rs("Reporting9") = Tsc79
                rs("Reporting10") = Tsc7A
                rs("Reporting11") = Tsc7B
                rs("Reporting12") = Tsc7C
                rs("Reporting13") = Tsc7D
                rs("Reporting14") = Tsc7E
                rs("Reporting15") = Tsc7F
                rs("statist1")  = Tsc81
                rs("statist2")  = Tsc82
                rs("statist3")  = Tsc83
                rs("Other1") = Tsc91
                rs("other2") = Tsc92
                rs("other3") = Tsc93
                rs("Other4") = Tsc94
                rs("other5") = Tsc95
                rs("other6") = Tsc96

 		rs.update
                end if
                rs.close
		conn.committrans
		msg = "紀錄已更新"
        end if
end if



%>
<html>
<head>
<title>用戶管理-新增</title>
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


function clearTs1(sidx){
       formObj=document.form1;
       if (sidx == 0 ){
          if ( formObj.Tsc1.checked==true){
             formObj.Tsc1.value = 1 ;
             formObj.Tsc11.value = 1 ;
             formObj.Tsc11.checked=true;
             formObj.Tsc12.value = 1 ;
             formObj.Tsc12.checked=true;
             formObj.Tsc13.value = 1 ;
             formObj.Tsc13.checked=true;
             formObj.Tsc14.value = 1 ;
             formObj.Tsc14.checked=true;
             formObj.Tsc15.value = 1 ;
             formObj.Tsc15.checked=true;
             formObj.Tsc16.value = 1 ;
             formObj.Tsc16.checked=true;
       }else{
             formObj.Tsc1.value = 0 ;
             formObj.Tsc11.value = 0 ;
             formObj.Tsc11.checked=false;
             formObj.Tsc12.value = 0 ;
             formObj.Tsc12.checked=false;
             formObj.Tsc13.value = 0 ;
             formObj.Tsc13.checked=false;
             formObj.Tsc14.value = 0 ;
             formObj.Tsc14.checked=false;
             formObj.Tsc15.value = 0 ;
             formObj.Tsc15.checked=false;
             formObj.Tsc16.value = 0 ;
             formObj.Tsc16.checked=false;
       } 
       }
       if (sidx == 1 ){
          if ( formObj.Tsc11.checked==true){
             formObj.Tsc11.value = 1 ;            
             formObj.Tsc11.checked=true; 
          }else{
             formObj.Tsc11.value = 0 ;  
          }
          }
       if (sidx == 2 ){
          if ( formObj.Tsc12.checked==true){
             formObj.Tsc12.value = 1 ;            
             formObj.Tsc12.checked=true; 
          }else{
             formObj.Tsc12.value = 0 ;  
          }
          } 
       if (sidx == 3 ){
          if ( formObj.Tsc13.checked==true){
             formObj.Tsc13.value = 1 ;            
             formObj.Tsc13.checked=true; 
          }else{
             formObj.Tsc13.value = 0 ;  
          }
          }  
       if (sidx == 4 ){
          if ( formObj.Tsc14.checked==true){
             formObj.Tsc14.value = 1 ;            
             formObj.Tsc14.checked=true; 
          }else{
             formObj.Tsc14.value = 0 ;  
          }
          }
       if (sidx == 5 ){
          if ( formObj.Tsc15.checked==true){
             formObj.Tsc15.value = 1 ;            
             formObj.Tsc15.checked=true; 
          }else{
             formObj.Tsc15.value = 0 ;  
          }
          }
       if (sidx == 6 ){
          if ( formObj.Tsc16.checked==true){
             formObj.Tsc16.value = 1 ;            
             formObj.Tsc16.checked=true; 
          }else{
             formObj.Tsc16.value = 0 ;  
          }
          }  
}
function clearTs2(sidx){
       formObj=document.form1;
       if (sidx == 0 ){
          if ( formObj.Tsc2.checked==true){
             formObj.Tsc2.value = 1 ;
             formObj.Tsc21.value = 1 ;
             formObj.Tsc21.checked=true;
             formObj.Tsc22.value = 1 ;
             formObj.Tsc22.checked=true;
             formObj.Tsc23.value = 1 ;
             formObj.Tsc23.checked=true;
             formObj.Tsc24.value = 1 ;
             formObj.Tsc24.checked=true;
             formObj.Tsc25.value = 1 ;
             formObj.Tsc25.checked=true;
             formObj.Tsc26.value = 1 ;
             formObj.Tsc26.checked=true;
            formObj.Tsc27.value = 1 ;
             formObj.Tsc27.checked=true;
             formObj.Tsc28.value = 1 ;
             formObj.Tsc28.checked=true;
             formObj.Tsc29.value = 1 ;
             formObj.Tsc29.checked=true;
             formObj.Tsc2A.value = 1 ;
             formObj.Tsc2A.checked=true;
             formObj.Tsc2B.value = 1 ;
             formObj.Tsc2B.checked=true;
             formObj.Tsc2C.value = 1 ;
             formObj.Tsc2C.checked=true;        
       
       }else{
             formObj.Tsc2.value = 0 ;
             formObj.Tsc21.value = 0 ;
             formObj.Tsc21.checked=false;
             formObj.Tsc22.value = 0 ;
             formObj.Tsc22.checked=false;
             formObj.Tsc23.value = 0 ;
             formObj.Tsc23.checked=false;
             formObj.Tsc24.value = 0 ;
             formObj.Tsc24.checked=false;
             formObj.Tsc25.value = 0 ;
             formObj.Tsc25.checked=false;
             formObj.Tsc26.value = 0 ;
             formObj.Tsc26.checked=false;
             formObj.Tsc27.value = 0 ;
             formObj.Tsc27.checked=false;
             formObj.Tsc28.value = 0 ;
             formObj.Tsc28.checked=false;
             formObj.Tsc29.value = 0 ;
             formObj.Tsc29.checked=false;
             formObj.Tsc2A.value = 0 ;
             formObj.Tsc2A.checked=false;
             formObj.Tsc2B.value = 0 ;
             formObj.Tsc2B.checked=false;
             formObj.Tsc2C.value = 0 ;
             formObj.Tsc2C.checked=false;         
           
       } 
       }
       if (sidx == 1 ){
          if ( formObj.Tsc21.checked==true){
             formObj.Tsc21.value = 1 ;            
             formObj.Tsc21.checked=true; 
          }else{
             formObj.Tsc21.value = 0 ;  
          }
          }
       if (sidx == 2 ){
          if ( formObj.Tsc22.checked==true){
             formObj.Tsc22.value = 1 ;            
             formObj.Tsc22.checked=true; 
          }else{
             formObj.Tsc22.value = 0 ;  
          }
          } 
       if (sidx == 3 ){
          if ( formObj.Tsc23.checked==true){
             formObj.Tsc23.value = 1 ;            
             formObj.Tsc23.checked=true; 
          }else{
             formObj.Tsc23.value = 0 ;  
          }
          }  
       if (sidx == 4 ){
          if ( formObj.Tsc24.checked==true){
             formObj.Tsc24.value = 1 ;            
             formObj.Tsc24.checked=true; 
          }else{
             formObj.Tsc24.value = 0 ;  
          }
          }
       if (sidx == 5 ){
          if ( formObj.Tsc25.checked==true){
             formObj.Tsc25.value = 1 ;            
             formObj.Tsc25.checked=true; 
          }else{
             formObj.Tsc25.value = 0 ;  
          }
          }
       if (sidx == 6 ){
          if ( formObj.Tsc26.checked==true){
             formObj.Tsc26.value = 1 ;            
             formObj.Tsc26.checked=true; 
          }else{
             formObj.Tsc26.value = 0 ;  
          }
          }  
       if (sidx == 7 ){
          if ( formObj.Tsc27.checked==true){
             formObj.Tsc27.value = 1 ;            
             formObj.Tsc27.checked=true; 
          }else{
             formObj.Tsc27.value = 0 ;  
          }
          }
       if (sidx == 8 ){
          if ( formObj.Tsc28.checked==true){
             formObj.Tsc28.value = 1 ;            
             formObj.Tsc28.checked=true; 
          }else{
             formObj.Tsc28.value = 0 ;  
          }
          } 
       if (sidx == 9 ){
          if ( formObj.Tsc29.checked==true){
             formObj.Tsc29.value = 1 ;            
             formObj.Tsc29.checked=true; 
          }else{
             formObj.Tsc29.value = 0 ;  
          }
          }  
       if (sidx == 10 ){
          if ( formObj.Tsc2A.checked==true){
             formObj.Tsc2A.value = 1 ;            
             formObj.Tsc2A.checked=true; 
          }else{
             formObj.Tsc2A.value = 0 ;  
          }
          }
       if (sidx == 11 ){
          if ( formObj.Tsc2B.checked==true){
             formObj.Tsc2B.value = 1 ;            
             formObj.Tsc2B.checked=true; 
          }else{
             formObj.Tsc2B.value = 0 ;  
          }
          }
       if (sidx == 12 ){
          if ( formObj.Tsc2C.checked==true){
             formObj.Tsc2C.value = 1 ;            
             formObj.Tsc2C.checked=true; 
          }else{
             formObj.Tsc2C.value = 0 ;  
          }
          }
}
function clearTs3(sidx){
       formObj=document.form1;
       if (sidx == 0 ){
          if ( formObj.Tsc3.checked==true){
             formObj.Tsc3.value = 1 ;
             formObj.Tsc31.value = 1 ;
             formObj.Tsc31.checked=true;
             formObj.Tsc32.value = 1 ;
             formObj.Tsc32.checked=true;
             formObj.Tsc33.value = 1 ;
             formObj.Tsc33.checked=true;
             formObj.Tsc34.value = 1 ;
             formObj.Tsc34.checked=true;
             formObj.Tsc35.value = 1 ;
             formObj.Tsc35.checked=true;
             formObj.Tsc36.value = 1 ;
             formObj.Tsc36.checked=true;
             formObj.Tsc37.value = 1 ;
             formObj.Tsc37.checked=true;
            formObj.Tsc38.value = 1 ;
             formObj.Tsc38.checked=true;  
         
       
       }else{
             formObj.Tsc3.value = 0 ;
             formObj.Tsc31.value = 0 ;
             formObj.Tsc31.checked=false;
             formObj.Tsc32.value = 0 ;
             formObj.Tsc32.checked=false;
             formObj.Tsc33.value = 0 ;
             formObj.Tsc33.checked=false;
             formObj.Tsc34.value = 0 ;
             formObj.Tsc34.checked=false;
             formObj.Tsc35.value = 0 ;
             formObj.Tsc35.checked=false;
             formObj.Tsc36.value = 0 ;
             formObj.Tsc36.checked=false;
             formObj.Tsc37.value = 0;
             formObj.Tsc37.checked=false;
            formObj.Tsc38.value = 0;
             formObj.Tsc38.checked=false; 
           
       } 
       }
       if (sidx == 1 ){
          if ( formObj.Tsc31.checked==true){
             formObj.Tsc31.value = 1 ;            
             formObj.Tsc31.checked=true; 
          }else{
             formObj.Tsc31.value = 0 ;  
          }
          }
       if (sidx == 2 ){
          if ( formObj.Tsc32.checked==true){
             formObj.Tsc32.value = 1 ;            
             formObj.Tsc32.checked=true; 
          }else{
             formObj.Tsc32.value = 0 ;  
          }
          } 
       if (sidx == 3 ){
          if ( formObj.Tsc33.checked==true){
             formObj.Tsc33.value = 1 ;            
             formObj.Tsc33.checked=true; 
          }else{
             formObj.Tsc33.value = 0 ;  
          }
          }  
       if (sidx == 4 ){
          if ( formObj.Tsc34.checked==true){
             formObj.Tsc34.value = 1 ;            
             formObj.Tsc34.checked=true; 
          }else{
             formObj.Tsc34.value = 0 ;  
          }
          }
       if (sidx == 5 ){
          if ( formObj.Tsc35.checked==true){
             formObj.Tsc35.value = 1 ;            
             formObj.Tsc35.checked=true; 
          }else{
             formObj.Tsc35.value = 0 ;  
          }
          }
       if (sidx == 6 ){
          if ( formObj.Tsc36.checked==true){
             formObj.Tsc36.value = 1 ;            
             formObj.Tsc36.checked=true; 
          }else{
             formObj.Tsc36.value = 0 ;  
          }
          }  
       if (sidx == 7 ){
          if ( formObj.Tsc37.checked==true){
             formObj.Tsc37.value = 1 ;            
             formObj.Tsc37.checked=true; 
          }else{
             formObj.Tsc37.value = 0 ;  
          }
          }
      if (sidx == 8 ){
          if ( formObj.Tsc38.checked==true){
             formObj.Tsc38.value = 1 ;            
             formObj.Tsc38.checked=true; 
          }else{
             formObj.Tsc38.value = 0 ;  
          }
          }
}
function clearTs4(sidx){
       formObj=document.form1;
       if (sidx == 0 ){
          if ( formObj.Tsc4.checked==true){
             formObj.Tsc4.value = 1 ;
             formObj.Tsc41.value = 1 ;
             formObj.Tsc41.checked=true;
             formObj.Tsc42.value = 1 ;
             formObj.Tsc42.checked=true;
             formObj.Tsc43.value = 1 ;
             formObj.Tsc43.checked=true;
             formObj.Tsc44.value = 1 ;
             formObj.Tsc44.checked=true;
             formObj.Tsc45.value = 1 ;
             formObj.Tsc45.checked=true;
             formObj.Tsc46.value = 1 ;
             formObj.Tsc46.checked=true;
            formObj.Tsc47.value = 1 ;
             formObj.Tsc47.checked=true;
             formObj.Tsc48.value = 1 ;
             formObj.Tsc48.checked=true;
             formObj.Tsc49.value = 1 ;
             formObj.Tsc49.checked=true;
             formObj.Tsc4A.value = 1 ;
             formObj.Tsc4A.checked=true;
             formObj.Tsc4B.value = 1 ;
             formObj.Tsc4B.checked=true;
             formObj.Tsc4C.value = 1 ;
             formObj.Tsc4C.checked=true;
             formObj.Tsc4D.value = 1 ;
             formObj.Tsc4D.checked=true;        
       
       }else{
             formObj.Tsc4.value = 0 ;
             formObj.Tsc41.value = 0 ;
             formObj.Tsc41.checked=false;
             formObj.Tsc42.value = 0 ;
             formObj.Tsc42.checked=false;
             formObj.Tsc43.value = 0 ;
             formObj.Tsc43.checked=false;
             formObj.Tsc44.value = 0 ;
             formObj.Tsc44.checked=false;
             formObj.Tsc45.value = 0 ;
             formObj.Tsc45.checked=false;
             formObj.Tsc46.value = 0 ;
             formObj.Tsc46.checked=false;
             formObj.Tsc47.value = 0 ;
             formObj.Tsc47.checked=false;
             formObj.Tsc48.value = 0 ;
             formObj.Tsc48.checked=false;
             formObj.Tsc49.value = 0 ;
             formObj.Tsc49.checked=false;
             formObj.Tsc4A.value = 0 ;
             formObj.Tsc4A.checked=false;
             formObj.Tsc4B.value = 0 ;
             formObj.Tsc4B.checked=false;
             formObj.Tsc4C.value = 0 ;
             formObj.Tsc4C.checked=false;
             formObj.Tsc4D.value = 0 ;
             formObj.Tsc4D.checked=false; 
       } 
       }
       if (sidx == 1 ){
          if ( formObj.Tsc41.checked==true){
             formObj.Tsc41.value = 1 ;            
             formObj.Tsc41.checked=true; 
          }else{
             formObj.Tsc41.value = 0 ;  
          }
          }
       if (sidx == 2 ){
          if ( formObj.Tsc42.checked==true){
             formObj.Tsc42.value = 1 ;            
             formObj.Tsc42.checked=true; 
          }else{
             formObj.Tsc42.value = 0 ;  
          }
          } 
       if (sidx == 3 ){
          if ( formObj.Tsc43.checked==true){
             formObj.Tsc43.value = 1 ;            
             formObj.Tsc43.checked=true; 
          }else{
             formObj.Tsc43.value = 0 ;  
          }
          }  
       if (sidx == 4 ){
          if ( formObj.Tsc44.checked==true){
             formObj.Tsc44.value = 1 ;            
             formObj.Tsc44.checked=true; 
          }else{
             formObj.Tsc44.value = 0 ;  
          }
          }
       if (sidx == 5 ){
          if ( formObj.Tsc45.checked==true){
             formObj.Tsc45.value = 1 ;            
             formObj.Tsc45.checked=true; 
          }else{
             formObj.Tsc45.value = 0 ;  
          }
          }
       if (sidx == 6 ){
          if ( formObj.Tsc46.checked==true){
             formObj.Tsc46.value = 1 ;            
             formObj.Tsc46.checked=true; 
          }else{
             formObj.Tsc46.value = 0 ;  
          }
          }  
       if (sidx == 7 ){
          if ( formObj.Tsc47.checked==true){
             formObj.Tsc47.value = 1 ;            
             formObj.Tsc47.checked=true; 
          }else{
             formObj.Tsc47.value = 0 ;  
          }
          }
       if (sidx == 8 ){
          if ( formObj.Tsc48.checked==true){
             formObj.Tsc48.value = 1 ;            
             formObj.Tsc48.checked=true; 
          }else{
             formObj.Tsc48.value = 0 ;  
          }
          } 
       if (sidx == 9 ){
          if ( formObj.Tsc49.checked==true){
             formObj.Tsc49.value = 1 ;            
             formObj.Tsc49.checked=true; 
          }else{
             formObj.Tsc49.value = 0 ;  
          }
          }  
       if (sidx == 10 ){
          if ( formObj.Tsc4A.checked==true){
             formObj.Tsc4A.value = 1 ;            
             formObj.Tsc4A.checked=true; 
          }else{
             formObj.Tsc4A.value = 0 ;  
          }
          }
       if (sidx == 11 ){
          if ( formObj.Tsc4B.checked==true){
             formObj.Tsc4B.value = 1 ;            
             formObj.Tsc4B.checked=true; 
          }else{
             formObj.Tsc4B.value = 0 ;  
          }
          }
       if (sidx == 12 ){
          if ( formObj.Tsc4C.checked==true){
             formObj.Tsc4C.value = 1 ;            
             formObj.Tsc4C.checked=true; 
          }else{
             formObj.Tsc4C.value = 0 ;  
          }
          }
       if (sidx == 13 ){
          if ( formObj.Tsc4D.checked==true){
             formObj.Tsc4D.value = 1 ;            
             formObj.Tsc4D.checked=true; 
          }else{
             formObj.Tsc4D.value = 0 ;  
          }
          }
}

function clearTs5(sidx){
       formObj=document.form1;
       if (sidx == 0 ){
          if ( formObj.Tsc5.checked==true){
             formObj.Tsc5.value = 1 ;
             formObj.Tsc51.value = 1 ;
             formObj.Tsc51.checked=true;
             formObj.Tsc52.value = 1 ;
             formObj.Tsc52.checked=true;
             formObj.Tsc53.value = 1 ;
             formObj.Tsc53.checked=true;
             formObj.Tsc54.value = 1 ;
             formObj.Tsc54.checked=true;
             formObj.Tsc55.value = 1 ;
             formObj.Tsc55.checked=true;
             formObj.Tsc56.value = 1 ;
             formObj.Tsc56.checked=true;
             formObj.Tsc57.value = 1 ;
             formObj.Tsc57.checked=true;
             formObj.Tsc58.value = 1 ;
             formObj.Tsc58.checked=true;
             formObj.Tsc59.value = 1 ;
             formObj.Tsc59.checked=true;
             formObj.Tsc5A.value = 1 ;
             formObj.Tsc5A.checked=true;
             formObj.Tsc5B.value = 1 ;
             formObj.Tsc5B.checked=true;
             formObj.Tsc5C.value = 1 ;
             formObj.Tsc5C.checked=true;         
       
       }else{
             formObj.Tsc5.value = 0 ;
             formObj.Tsc51.value = 0 ;
             formObj.Tsc51.checked=false;
             formObj.Tsc52.value = 0 ;
             formObj.Tsc52.checked=false;
             formObj.Tsc53.value = 0 ;
             formObj.Tsc53.checked=false;
             formObj.Tsc54.value = 0 ;
             formObj.Tsc54.checked=false;
             formObj.Tsc55.value = 0 ;
             formObj.Tsc55.checked=false;
             formObj.Tsc56.value = 0 ;
             formObj.Tsc56.checked=false;
             formObj.Tsc57.value = 0 ;
             formObj.Tsc57.checked=false;
             formObj.Tsc58.value = 0 ;
             formObj.Tsc58.checked=false;
             formObj.Tsc59.value = 0 ;
             formObj.Tsc59.checked=false;
             formObj.Tsc5A.value = 0 ;
             formObj.Tsc5A.checked=false;
             formObj.Tsc5B.value = 0 ;
             formObj.Tsc5B.checked=false;
             formObj.Tsc5C.value = 0 ;
             formObj.Tsc5C.checked=false;       
       } 
       }
       if (sidx == 1 ){
          if ( formObj.Tsc51.checked==true){
             formObj.Tsc51.value = 1 ;            
             formObj.Tsc51.checked=true; 
          }else{
             formObj.Tsc51.value = 0 ;  
          }
          }
       if (sidx == 2 ){
          if ( formObj.Tsc52.checked==true){
             formObj.Tsc52.value = 1 ;            
             formObj.Tsc52.checked=true; 
          }else{
             formObj.Tsc52.value = 0 ;  
          }
          } 
       if (sidx == 3 ){
          if ( formObj.Tsc53.checked==true){
             formObj.Tsc53.value = 1 ;            
             formObj.Tsc53.checked=true; 
          }else{
             formObj.Tsc53.value = 0 ;  
          }
          }  
       if (sidx == 4 ){
          if ( formObj.Tsc54.checked==true){
             formObj.Tsc54.value = 1 ;            
             formObj.Tsc54.checked=true; 
          }else{
             formObj.Tsc54.value = 0 ;  
          }
          }
       if (sidx == 5 ){
          if ( formObj.Tsc55.checked==true){
             formObj.Tsc55.value = 1 ;            
             formObj.Tsc55.checked=true; 
          }else{
             formObj.Tsc55.value = 0 ;  
          }
          }
       if (sidx == 6 ){
          if ( formObj.Tsc56.checked==true){
             formObj.Tsc56.value = 1 ;            
             formObj.Tsc56.checked=true; 
          }else{
             formObj.Tsc56.value = 0 ;  
          }
          }  
       if (sidx == 7 ){
          if ( formObj.Tsc57.checked==true){
             formObj.Tsc57.value = 1 ;            
             formObj.Tsc57.checked=true; 
          }else{
             formObj.Tsc57.value = 0 ;  
          }
          }
       if (sidx == 8 ){
          if ( formObj.Tsc58.checked==true){
             formObj.Tsc58.value = 1 ;            
             formObj.Tsc58.checked=true; 
          }else{
             formObj.Tsc58.value = 0 ;  
          }
          } 
       if (sidx == 9 ){
          if ( formObj.Tsc59.checked==true){
             formObj.Tsc59.value = 1 ;            
             formObj.Tsc59.checked=true; 
          }else{
             formObj.Tsc59.value = 0 ;  
          }
          }  
       if (sidx == 10 ){
          if ( formObj.Tsc5A.checked==true){
             formObj.Tsc5A.value = 1 ;            
             formObj.Tsc5A.checked=true; 
          }else{
             formObj.Tsc5A.value = 0 ;  
          }
          }
       if (sidx == 11 ){
          if ( formObj.Tsc5B.checked==true){
             formObj.Tsc5B.value = 1 ;            
             formObj.Tsc5B.checked=true; 
          }else{
             formObj.Tsc5B.value = 0 ;  
          }
          }
      if (sidx == 12 ){
          if ( formObj.Tsc5C.checked==true){
             formObj.Tsc5C.value = 1 ;            
             formObj.Tsc5C.checked=true; 
          }else{
             formObj.Tsc5C.value = 0 ;  
          }
          }

}

function clearTs6(sidx){
       formObj=document.form1;
       if (sidx == 0 ){
          if ( formObj.Tsc6.checked==true){
             formObj.Tsc6.value = 1 ;
             formObj.Tsc61.value = 1 ;
             formObj.Tsc61.checked=true;

       }else{
             formObj.Tsc6.value = 0 ;
             formObj.Tsc61.value = 0 ;
             formObj.Tsc61.checked=false;

       } 
       }
       if (sidx == 1 ){
          if ( formObj.Tsc61.checked==true){
             formObj.Tsc61.value = 1 ;            
             formObj.Tsc61.checked=true; 
          }else{
             formObj.Tsc61.value = 0 ;  
          }
          }
}

function clearTs7(sidx){
       formObj=document.form1;
       if (sidx == 0 ){
          if ( formObj.Tsc7.checked==true){
             formObj.Tsc7.value = 1 ;
             formObj.Tsc71.value = 1 ;
             formObj.Tsc71.checked=true;
             formObj.Tsc72.value = 1 ;
             formObj.Tsc72.checked=true;
             formObj.Tsc73.value = 1 ;
             formObj.Tsc73.checked=true;
             formObj.Tsc74.value = 1 ;
             formObj.Tsc74.checked=true;
             formObj.Tsc75.value = 1 ;
             formObj.Tsc75.checked=true;
             formObj.Tsc76.value = 1 ;
             formObj.Tsc76.checked=true;
             formObj.Tsc77.value = 1 ;
             formObj.Tsc77.checked=true;
             formObj.Tsc78.value = 1 ;
             formObj.Tsc78.checked=true;
             formObj.Tsc79.value = 1 ;
             formObj.Tsc79.checked=true;
             formObj.Tsc7A.value = 1 ;
             formObj.Tsc7A.checked=true;
             formObj.Tsc7B.value = 1 ;
             formObj.Tsc7B.checked=true;
             formObj.Tsc7C.value = 1 ;
             formObj.Tsc7C.checked=true;         
             formObj.Tsc7D.value = 1 ;
             formObj.Tsc7D.checked=true;
             formObj.Tsc7E.value = 1 ;
             formObj.Tsc7E.checked=true;             
             formObj.Tsc7F.value = 1 ;
             formObj.Tsc7F.checked=true;             
       }else{
             formObj.Tsc7.value = 0 ;
             formObj.Tsc71.value = 0 ;
             formObj.Tsc71.checked=false;
             formObj.Tsc72.value = 0 ;
             formObj.Tsc72.checked=false;
             formObj.Tsc73.value = 0 ;
             formObj.Tsc73.checked=false;
             formObj.Tsc74.value = 0 ;
             formObj.Tsc74.checked=false;
             formObj.Tsc75.value = 0 ;
             formObj.Tsc75.checked=false;
             formObj.Tsc76.value = 0 ;
             formObj.Tsc76.checked=false;
             formObj.Tsc77.value = 0 ;
             formObj.Tsc77.checked=false;
             formObj.Tsc78.value = 0 ;
             formObj.Tsc78.checked=false;
             formObj.Tsc79.value = 0 ;
             formObj.Tsc79.checked=false;
             formObj.Tsc7A.value = 0 ;
             formObj.Tsc7A.checked=false;
             formObj.Tsc7B.value = 0 ;
             formObj.Tsc7B.checked=false;
             formObj.Tsc7C.value = 0 ;
             formObj.Tsc7C.checked=false;       
            formObj.Tsc7D.value = 0 ;
             formObj.Tsc7D.checked=false;
             formObj.Tsc7E.value = 0 ;
             formObj.Tsc7E.checked=false;  
             formObj.Tsc7F.value = 0 ;
             formObj.Tsc7F.checked=false;        
       } 
       }
       if (sidx == 1 ){
          if ( formObj.Tsc71.checked==true){
             formObj.Tsc71.value = 1 ;            
             formObj.Tsc71.checked=true; 
          }else{
             formObj.Tsc71.value = 0 ;  
          }
          }
       if (sidx == 2 ){
          if ( formObj.Tsc72.checked==true){
             formObj.Tsc72.value = 1 ;            
             formObj.Tsc72.checked=true; 
          }else{
             formObj.Tsc72.value = 0 ;  
          }
          } 
       if (sidx == 3 ){
          if ( formObj.Tsc73.checked==true){
             formObj.Tsc73.value = 1 ;            
             formObj.Tsc73.checked=true; 
          }else{
             formObj.Tsc73.value = 0 ;  
          }
          }  
       if (sidx == 4 ){
          if ( formObj.Tsc74.checked==true){
             formObj.Tsc74.value = 1 ;            
             formObj.Tsc74.checked=true; 
          }else{
             formObj.Tsc74.value = 0 ;  
          }
          }
       if (sidx == 5 ){
          if ( formObj.Tsc75.checked==true){
             formObj.Tsc75.value = 1 ;            
             formObj.Tsc75.checked=true; 
          }else{
             formObj.Tsc75.value = 0 ;  
          }
          }
       if (sidx == 6 ){
          if ( formObj.Tsc76.checked==true){
             formObj.Tsc76.value = 1 ;            
             formObj.Tsc76.checked=true; 
          }else{
             formObj.Tsc76.value = 0 ;  
          }
          }  
       if (sidx == 7 ){
          if ( formObj.Tsc77.checked==true){
             formObj.Tsc77.value = 1 ;            
             formObj.Tsc77.checked=true; 
          }else{
             formObj.Tsc77.value = 0 ;  
          }
          }
       if (sidx == 8 ){
          if ( formObj.Tsc78.checked==true){
             formObj.Tsc78.value = 1 ;            
             formObj.Tsc78.checked=true; 
          }else{
             formObj.Tsc78.value = 0 ;  
          }
          } 
       if (sidx == 9 ){
          if ( formObj.Tsc79.checked==true){
             formObj.Tsc79.value = 1 ;            
             formObj.Tsc79.checked=true; 
          }else{
             formObj.Tsc79.value = 0 ;  
          }
          }  
       if (sidx == 10 ){
          if ( formObj.Tsc7A.checked==true){
             formObj.Tsc7A.value = 1 ;            
             formObj.Tsc7A.checked=true; 
          }else{
             formObj.Tsc7A.value = 0 ;  
          }
          }
       if (sidx == 11 ){
          if ( formObj.Tsc7B.checked==true){
             formObj.Tsc7B.value = 1 ;            
             formObj.Tsc7B.checked=true; 
          }else{
             formObj.Tsc7B.value = 0 ;  
          }
          }
      if (sidx == 12 ){
          if ( formObj.Tsc7C.checked==true){
             formObj.Tsc7C.value = 1 ;            
             formObj.Tsc7C.checked=true; 
          }else{
             formObj.Tsc7C.value = 0 ;  
          }
          }
       if (sidx == 13 ){
          if ( formObj.Tsc7D.checked==true){
             formObj.Tsc7D.value = 1 ;            
             formObj.Tsc7D.checked=true; 
          }else{
             formObj.Tsc7D.value = 0 ;  
          }
          }
      if (sidx == 14 ){
          if ( formObj.Tsc7E.checked==true){
             formObj.Tsc7E.value = 1 ;            
             formObj.Tsc7E.checked=true; 
          }else{
             formObj.Tsc7E.value = 0 ;  
          }
          }
      if (sidx == 15 ){
          if ( formObj.Tsc7F.checked==true){
             formObj.Tsc7F.value = 1 ;            
             formObj.Tsc7F.checked=true; 
          }else{
             formObj.Tsc7F.value = 0 ;  
          }
          }

}

function clearTs8(sidx){
       formObj=document.form1;
       if (sidx == 0 ){
          if ( formObj.Tsc8.checked==true){
             formObj.Tsc8.value = 1 ;
             formObj.Tsc81.value = 1 ;
             formObj.Tsc81.checked=true;
             formObj.Tsc82.value = 1 ;
             formObj.Tsc82.checked=true;
             formObj.Tsc83.value = 1 ;
             formObj.Tsc83.checked=true;

       }else{
             formObj.Tsc8.value = 0 ;
             formObj.Tsc81.value = 0 ;
             formObj.Tsc81.checked=false;
             formObj.Tsc82.value = 0 ;
             formObj.Tsc82.checked=false;
             formObj.Tsc83.value = 0 ;
             formObj.Tsc83.checked=false;

       } 
       }
       if (sidx == 1 ){
          if ( formObj.Tsc81.checked==true){
             formObj.Tsc81.value = 1 ;            
             formObj.Tsc81.checked=true; 
          }else{
             formObj.Tsc81.value = 0 ;  
          }
          }
       if (sidx == 2 ){
          if ( formObj.Tsc82.checked==true){
             formObj.Tsc82.value = 1 ;            
             formObj.Tsc82.checked=true; 
          }else{
             formObj.Tsc82.value = 0 ;  
          }
          } 
       if (sidx == 3 ){
          if ( formObj.Tsc83.checked==true){
             formObj.Tsc83.value = 1 ;            
             formObj.Tsc83.checked=true; 
          }else{
             formObj.Tsc83.value = 0 ;  
          }
          }  


}

function clearTs9(sidx){
       formObj=document.form1;
       if (sidx == 0 ){
          if ( formObj.Tsc9.checked==true){
             formObj.Tsc9.value = 1 ;
             formObj.Tsc91.value = 1 ;
             formObj.Tsc91.checked=true;
             formObj.Tsc92.value = 1 ;
             formObj.Tsc92.checked=true;
             formObj.Tsc93.value = 1 ;
             formObj.Tsc93.checked=true;
             formObj.Tsc94.value = 1 ;
             formObj.Tsc94.checked=true;
             formObj.Tsc95.value = 1 ;
             formObj.Tsc95.checked=true;
             formObj.Tsc96.value = 1 ;
             formObj.Tsc96.checked=true;

       }else{
             formObj.Tsc9.value = 0 ;
             formObj.Tsc91.value = 0 ;
             formObj.Tsc91.checked=false;
             formObj.Tsc92.value = 0 ;
             formObj.Tsc92.checked=false;
             formObj.Tsc93.value = 0 ;
             formObj.Tsc93.checked=false;
             formObj.Tsc94.value = 0 ;
             formObj.Tsc94.checked=false;
             formObj.Tsc95.value = 0 ;
             formObj.Tsc95.checked=false;
             formObj.Tsc96.value = 0 ;
             formObj.Tsc96.checked=false;

       } 
       }
       if (sidx == 1 ){
          if ( formObj.Tsc91.checked==true){
             formObj.Tsc91.value = 1 ;            
             formObj.Tsc91.checked=true; 
          }else{
             formObj.Tsc91.value = 0 ;  
          }
          }
       if (sidx == 2 ){
          if ( formObj.Tsc92.checked==true){
             formObj.Tsc92.value = 1 ;            
             formObj.Tsc92.checked=true; 
          }else{
             formObj.Tsc92.value = 0 ;  
          }
          } 
       if (sidx == 3 ){
          if ( formObj.Tsc93.checked==true){
             formObj.Tsc93.value = 1 ;            
             formObj.Tsc93.checked=true; 
          }else{
             formObj.Tsc93.value = 0 ;  
          }
          }  
       if (sidx == 4 ){
          if ( formObj.Tsc94.checked==true){
             formObj.Tsc94.value = 1 ;            
             formObj.Tsc94.checked=true; 
          }else{
             formObj.Tsc94.value = 0 ;  
          }
          }
       if (sidx == 5 ){
          if ( formObj.Tsc95.checked==true){
             formObj.Tsc95.value = 1 ;            
             formObj.Tsc95.checked=true; 
          }else{
             formObj.Tsc95.value = 0 ;  
          }
          }
       if (sidx == 6 ){
          if ( formObj.Tsc96.checked==true){
             formObj.Tsc96.value = 1 ;            
             formObj.Tsc96.checked=true; 
          }else{
             formObj.Tsc96.value = 0 ;  
          }
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

<%if id="" then%>
	if (formObj.password.value==""){
		reqField=reqField+", 密碼";
		if (!placeFocus)
			placeFocus=formObj.password;
	}
<%end if%>

	if (formObj.password.value!=formObj.password1.value){
		reqField=reqField+", 相符的重入密碼";
		if (!placeFocus)
			placeFocus=formObj.password;
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
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0">
<!-- #include file="menu.asp" -->
<%if msg<>"" then %>
<div><center><font size="3"><%=msg%></font></center></div>
<% end if%>

<br>
<center>
<form name="form1" method="post" action="userAdd.asp">
<div><center><font size="3">用戶管理-新增</font></center></div>
<br>
<table border="0" cellpadding="0" cellspacing="0">
    <tr>
        <td width="130" align="right">使用者名稱</td>
        <td width="10">&nbsp</td>
        <td width="170" ><input type="text" name="username" value ="<%=username%>" size="20"></td>
        <td width="100" >&nbsp;</td>
    </tr>
    <tr>
         <td width="130" align="right">密碼</b></td>
        <td width="10">&nbsp</td>
        <td><input type="password" name="password" size="20"></td>
        <td width="100">&nbsp;</td>
    </tr>
    <tr>
	 <td width="130" align="right">重入密碼</td>	
        <td width="10">&nbsp</td>
	<td><input type="password" name="password1" value="<%=password1%>" size="20" ></td>
        <td width="100">&nbsp;</td>
   </tr>
   <tr>
	 <td width="130" align="right">用戶級別</td>
        <td width="10">&nbsp</td>
	<td><select name="userLevel">
<%for idx = 0 to 6%>
			<option value=<%=idx%> <%if idx=userLevel then response.write " selected" end if%>><%=arrLevel(idx)%></option>
<%next%>
			</select>
		</td>
          <td width="100">&nbsp;</td>
	</tr>
</table>
<br>

<table border="0" cellpadding="0" cellspacing="0" >
    <tr>
          <td  width="220"  valign="top">
               <table border="0" cellpadding="0"  >
                  <tr> 
                    <td><ul><input type="checkbox" name="Tsc1"  value="<%=Tsc1%>"  <%if Tsc1<>0 then response.write " checked" end if%>  onclick="clearTs1(0)">社員資料
                        <dd><input type="checkbox" name="Tsc11" value="<%=Tsc11%>" <%if Tsc11<>0 then response.write " checked" end if%> onclick="{clearTs1(1)}">加入新社員</DD>    
                        <dd><input type="checkbox" name="Tsc12" value="<%=Tsc12%>" <%if Tsc12<>0 then response.write " checked" end if%> onclick="{clearTs1(2)}">社員資料修正</DD>
                        <dd><input type="checkbox" name="Tsc13" value="<%=Tsc13%>" <%if Tsc13<>0 then response.write " checked" end if%> onclick="{clearTs1(3)}">轉換聯絡人建立</DD>    
                        <dd><input type="checkbox" name="Tsc14" value="<%=Tsc14%>" <%if Tsc14<>0 then response.write " checked" end if%> onclick="{clearTs1(4)}">銀行資料操作</DD>
                        <DD><input type="checkbox" name="Tsc15" value="<%=Tsc15%>" <%if Tsc15<>0 then response.write " checked" end if%> onclick="{clearTs1(5)}">新社員開戶建立</DD>    
                        <DD><input type="checkbox" name="Tsc16" value="<%=Tsc16%>" <%if Tsc16<>0 then response.write " checked" end if%> onclick="{clearTs1(6)}">截數設定建立</DD>
                        <DD>&nbsp</DD>
                        <DD>&nbsp</DD>
                        <DD>&nbsp</DD>
                        <DD>&nbsp</DD>
                        <DD>&nbsp</DD>
                        <DD>&nbsp</DD>
                        <DD>&nbsp</DD>
                        <DD>&nbsp</DD>
                        <DD>&nbsp</DD> 
                        <DD>&nbsp</DD>
                        </Ul></td>
                  </tr> 
                  <tr> 
                     <td><ul><input type="checkbox" name="Tsc5"  value="<%=Tsc5%>"  <%if Tsc5<>0 then response.write " checked" end if%>  onclick="clearTs5(0)">股金
                             <DD><input type="checkbox" name="Tsc51" value="<%=Tsc51%>" <%if Tsc51<>0 then response.write " checked" end if%>    onclick="{clearTs5(1)}">股息計算操作</DD>    
                             <DD><input type="checkbox" name="Tsc52" value="<%=Tsc52%>" <%if Tsc52<>0 then response.write " checked" end if%>    onclick="{clearTs5(2)}">股息列印</DD>
                             <DD><input type="checkbox" name="Tsc53" value="<%=Tsc53%>" <%if Tsc53<>0 then response.write " checked" end if%>    onclick="{clearTs5(3)}">派息分配建立</DD>    
                             <DD><input type="checkbox" name="Tsc54" value="<%=Tsc54%>" <%if Tsc54<>0 then response.write " checked" end if%>    onclick="{clearTs5(4)}">派息分配修改操作</DD>
                             <DD><input type="checkbox" name="Tsc55" value="<%=Tsc55%>" <%if Tsc55<>0 then response.write " checked" end if%>    onclick="{clearTs5(5)}">銀行派息磁碟建立</DD>    
                             <DD><input type="checkbox" name="Tsc56" value="<%=Tsc56%>" <%if Tsc56<>0 then response.write " checked" end if%>    onclick="{clearTs5(6)}">派息過數</DD>
                             <DD><input type="checkbox" name="Tsc57" value="<%=Tsc57%>" <%if Tsc57<>0 then response.write " checked" end if%>    onclick="{clearTs5(7)}">銀行轉帳失效建立</DD>    
                             <DD><input type="checkbox" name="Tsc58" value="<%=Tsc58%>" <%if Tsc58<>0 then response.write " checked" end if%>    onclick="{clearTs5(8)}">暫停派息過數</DD>
                             <DD><input type="checkbox" name="Tsc59" value="<%=Tsc59%>" <%if Tsc59<>0 then response.write " checked" end if%>    onclick="{clearTs5(9)}">退股建立</DD>    
                             <DD><input type="checkbox" name="Tsc5A" value="<%=Tsc5A%>" <%if Tsc5A<>0 then response.write " checked" end if%>    onclick="{clearTs5(10)}">現金存款建立</DD>
                             <DD><input type="checkbox" name="Tsc5B" value="<%=Tsc5B%>" <%if Tsc5B<>0 then response.write " checked" end if%>    onclick="{clearTs5(11)}">股金列印</DD>    
                             <DD><input type="checkbox" name="Tsc5C" value="<%=Tsc5C%>" <%if Tsc5C<>0 then response.write " checked" end if%>    onclick="{clearTs5(12)}">股金細項修正</DD>
                             <DD>&nbsp</DD>
                              <DD>&nbsp</DD> 
                            </Ul></td>
                  </tr> 
                </table>
            </td>
          
          <td  width="230"  valign="top">               
             <table border="0" cellpadding="0"  >
                <tr> 
                 <td><ul><input type="checkbox" name="Tsc2" value="<%=Tsc2%>" <%if Tsc2<>0 then response.write " checked" end if%>    onclick="{clearTs2(0)}">貸款
                             <DD><input type="checkbox" name="Tsc21" value="<%=Tsc21%>" <%if Tsc21<>0 then response.write " checked" end if%>    onclick="{clearTs2(1)}">貸款申請</DD>    
                             <DD><input type="checkbox" name="Tsc22" value="<%=Tsc22%>" <%if Tsc22<>0 then response.write " checked" end if%>    onclick="{clearTs2(2)}">新貸款建立</DD>
                             <DD><input type="checkbox" name="Tsc23" value="<%=Tsc23%>" <%if Tsc23<>0 then response.write " checked" end if%>    onclick="{clearTs2(3)}">貸款修正</DD>    
                             <DD><input type="checkbox" name="Tsc24" value="<%=Tsc24%>" <%if Tsc24<>0 then response.write " checked" end if%>    onclick="{clearTs2(4)}">貸款列印</DD>
                             <DD><input type="checkbox" name="Tsc25" value="<%=Tsc25%>" <%if Tsc25<>0 then response.write " checked" end if%>    onclick="{clearTs2(5)}">延期操作</DD>    
                             <DD><input type="checkbox" name="Tsc26" value="<%=Tsc26%>" <%if Tsc26<>0 then response.write " checked" end if%>    onclick="{clearTs2(6)}">現金還款</DD>
                             <DD><input type="checkbox" name="Tsc27" value="<%=Tsc27%>" <%if Tsc27<>0 then response.write " checked" end if%>    onclick="{clearT2(7)}">股金還款</DD>    
                             <DD><input type="checkbox" name="Tsc28" value="<%=Tsc28%>" <%if Tsc28<>0 then response.write " checked" end if%>    onclick="{clearTs2(8)}">貸款退款至股金操作</DD>
                             <DD><input type="checkbox" name="Tsc29" value="<%=Tsc29%>" <%if Tsc29<>0 then response.write " checked" end if%>    onclick="{clearTs2(9)}">劃消貸款建立</DD>    
                             <DD><input type="checkbox" name="Tsc2A" value="<%=Tsc2A%>" <%if Tsc2A<>0 then response.write " checked" end if%>    onclick="{clearTs2(10)}">貸款細項列印</DD>
                             <DD><input type="checkbox" name="Tsc2B" value="<%=Tsc2B%>" <%if Tsc2B<>0 then response.write " checked" end if%>    onclick="{clearTs2(11)}">貸款細項修正</DD>
                             <DD><input type="checkbox" name="Tsc2C" value="<%=Tsc2C%>" <%if Tsc2C<>0 then response.write " checked" end if%>    onclick="{clearTs2(12)}">取消銀行脫期建立</DD>
                             <DD>&nbsp</DD>           
                             <DD>&nbsp</DD>
                             </Ul></td>
                </tr>
                <tr> 
                    <td><ul><input type="checkbox" name="Tsc6"  value="<%=Tsc6%>"  <%if Tsc6<>0 then response.write " checked" end if%>  onclick="clearTs6(0)">個人戶口
                             <DD><input type="checkbox" name="Tsc61" value="<%=Tsc61%>" <%if Tsc61<>0 then response.write " checked" end if%>    onclick="{clearTs6(1)}">社員資料查詢請</DD>    
                             <DD>&nbsp</DD>
                             <DD>&nbsp</DD>
                             <DD>&nbsp</DD>
                             <DD>&nbsp</DD>
                             <DD>&nbsp</DD>
                             <DD>&nbsp</DD>
                             <DD>&nbsp</DD>
                             <DD>&nbsp</DD>
                             <DD>&nbsp</DD>
                             <DD>&nbsp</DD>
                             <DD>&nbsp</DD>
                             <DD>&nbsp</DD>
                             <DD>&nbsp</DD>
                            </Ul></td>
                </tr>
             </table>
         </td>
          <td  width="230"  valign="top">
             <table border="0" cellpadding="0"  >
                 <tr>
                   <td ><ul><input type="checkbox" name="Tsc3" value="<%=Tsc3%>" <%if Tsc3<>0 then response.write " checked" end if%>    onclick="{clearTs3(0);}">清數及破產操作
                             <DD><input type="checkbox" name="Tsc31" value="<%=Tsc31%>" <%if Tsc31<>0 then response.write " checked" end if%>    onclick="{clearTs3(1)}">循環貸款</DD>
                             <DD><input type="checkbox" name="Tsc32" value="<%=Tsc32%>" <%if Tsc32<>0 then response.write " checked" end if%>    onclick="{clearTs3(2)}">現金清數</DD>    
                             <DD><input type="checkbox" name="Tsc33" value="<%=Tsc33%>" <%if Tsc33<>0 then response.write " checked" end if%>    onclick="{clearTs3(3)}">股金清數</DD>
                             <DD><input type="checkbox" name="Tsc34" value="<%=Tsc34%>" <%if Tsc34<>0 then response.write " checked" end if%>    onclick="{clearTs3(4)}">現金清數(本金)</DD>    
                             <DD><input type="checkbox" name="Tsc35" value="<%=Tsc35%>" <%if Tsc35<>0 then response.write " checked" end if%>    onclick="{clearTs3(5)}">破產操作建立</DD>
                             <DD><input type="checkbox" name="Tsc36" value="<%=Tsc36%>" <%if Tsc36<>0 then response.write " checked" end if%>    onclick="{clearTs3(6)}">破產列印</DD>    
                             <DD><input type="checkbox" name="Tsc37" value="<%=Tsc37%>" <%if Tsc37<>0 then response.write " checked" end if%>    onclick="{clearTs3(7)}">IVA操作建立</DD>
                             <DD><input type="checkbox" name="Tsc38" value="<%=Tsc38%>" <%if Tsc38<>0 then response.write " checked" end if%>     onclick="{clearTs3(8)}">IVA列印</DD>    
                             <DD>&nbsp</DD>
                             <DD>&nbsp</DD>
                             <DD>&nbsp</DD>
                             <DD>&nbsp</DD>
                             <DD>&nbsp</DD>
                             <dd>&nbsp</DD> 
                             <DD>&nbsp</DD>
                             <DD>&nbsp</DD>
                            </Ul></td>
                 </tr>
                 <tr>
                         <td ><ul><input type="checkbox" name="Tsc7" value="<%=Tsc7%>" <%if Tsc7<>0 then response.write " checked" end if%>    onclick="{clearTs7(0)}">報表

                             <DD><input type="checkbox" name="Tsc71" value="<%=Tsc71%>" <%if Tsc71<>0 then response.write " checked" end if%>    onclick="{clearTs7(1)}">個人資料列表</DD>
                             <DD><input type="checkbox" name="Tsc72" value="<%=Tsc72%>" <%if Tsc72<>0 then response.write " checked" end if%>    onclick="{clearTs7(2)}">呆賬報告</DD>    
                             <DD><input type="checkbox" name="Tsc73" value="<%=Tsc73%>" <%if Tsc73<>0 then response.write " checked" end if%>    onclick="{clearTs7(3)}">冷戶報告</DD>
                             <DD><input type="checkbox" name="Tsc74" value="<%=Tsc74%>" <%if Tsc74<>0 then response.write " checked" end if%>    onclick="{clearTs7(4)}">IVA報告</DD>
                             <DD><input type="checkbox" name="Tsc75" value="<%=Tsc75%>" <%if Tsc75<>0 then response.write " checked" end if%>    onclick="{clearTs7(5)}">破產報告</DD>
                             <DD><input type="checkbox" name="Tsc76" value="<%=Tsc76%>" <%if Tsc76<>0 then response.write " checked" end if%>    onclick="{clearTs7(6)}">社員分組/組員列表</DD>    
                             <DD><input type="checkbox" name="Tsc77" value="<%=Tsc77%>" <%if Tsc77<>0 then response.write " checked" end if%>    onclick="{clearTs7(7)}">社員轉帳資料列表</DD>
                             <DD><input type="checkbox" name="Tsc78" value="<%=Tsc78%>" <%if Tsc78<>0 then response.write " checked" end if%>    onclick="{clearTs7(8)}">社員生日名單</DD>
                             <DD><input type="checkbox" name="Tsc79" value="<%=Tsc79%>" <%if Tsc79<>0 then response.write " checked" end if%>    onclick="{clearTs7(9)}">退休社員報告</DD>                           
                             <DD><input type="checkbox" name="Tsc7A" value="<%=Tsc7A%>" <%if Tsc7A<>0 then response.write " checked" end if%>    onclick="{clearTs7(10)}">社員狀況列印</DD>    
                             <DD><input type="checkbox" name="Tsc7B" value="<%=Tsc7B%>" <%if Tsc7B<>0 then response.write " checked" end if%>    onclick="{clearTs7(11)}">現金帳列表</DD>
                             <DD><input type="checkbox" name="Tsc7C" value="<%=Tsc7C%>" <%if Tsc7C<>0 then response.write " checked" end if%>    onclick="{clearTs7(12)}">銀行帳列表</DD>
                             <DD><input type="checkbox" name="Tsc7D" value="<%=Tsc7D%>" <%if Tsc7D<>0 then response.write " checked" end if%>    onclick="{clearTs7(13)}">其他帳列</DD>
                             <DD><input type="checkbox" name="Tsc7E" value="<%=Tsc7E%>" <%if Tsc7E<>0 then response.write " checked" end if%>    onclick="{clearTs7(14)}">每月帳統計列表</DD> 
                             <DD><input type="checkbox" name="Tsc7F" value="<%=Tsc7F%>" <%if Tsc7F<>0 then response.write " checked" end if%>    onclick="{clearTs7(15)}">半年結</DD> 
                           </Ul></td>
                 </tr>       
            </table>
         </td>
       <td  width="300"  valign="top">
             <table border="0" cellpadding="0"  >
                 <tr>                                
                       <td width="300" ><ul><input type="checkbox" name="Tsc4" value="<%=Tsc4%>" <%if Tsc4<>0 then response.write " checked" end if%>    onclick="{clearTs4(0)}">自動轉帳

                             <DD><input type="checkbox" name="Tsc41" value="<%=Tsc41%>" <%if Tsc41<>0 then response.write " checked" end if%>    onclick="{clearTs4(1)}">轉帳建立表</DD>
                             <DD><input type="checkbox" name="Tsc42" value="<%=Tsc42%>" <%if Tsc42<>0 then response.write " checked" end if%>    onclick="{clearTs4(2)}">特別個案轉帳輸入操作</DD>    
                             <DD><input type="checkbox" name="Tsc43" value="<%=Tsc43%>" <%if Tsc43<>0 then response.write " checked" end if%>    onclick="{clearTs4(3)}">銀行轉帳試算</DD>
                             <DD><input type="checkbox" name="Tsc44" value="<%=Tsc44%>" <%if Tsc44<>0 then response.write " checked" end if%>    onclick="{clearTs4(4)}">特別個案轉帳試算</DD>
                             <DD><input type="checkbox" name="Tsc45" value="<%=Tsc45%>" <%if Tsc45<>0 then response.write " checked" end if%>    onclick="{clearTs4(5)}">銀行轉帳磁碟建立</DD>
                             <DD><input type="checkbox" name="Tsc46" value="<%=Tsc46%>" <%if Tsc46<>0 then response.write " checked" end if%>    onclick="{clearTs4(6)}">銀行脫期建立</DD>    
                             <DD><input type="checkbox" name="Tsc47" value="<%=Tsc47%>" <%if Tsc47<>0 then response.write " checked" end if%>    onclick="{clearTs4(7)}">銀行轉帳過數</DD>
                             <DD><input type="checkbox" name="Tsc48" value="<%=Tsc48%>" <%if Tsc48<>0 then response.write " checked" end if%>    onclick="{clearTs4(8)}">銀行自動轉帳失效通知書建立</DD>
                             <DD><input type="checkbox" name="Tsc49" value="<%=Tsc49%>" <%if Tsc49<>0 then response.write " checked" end if%>    onclick="{clearTs4(9)}">銀行自動轉帳失效通知書列印</DD>
                             <DD><input type="checkbox" name="Tsc4A" value="<%=Tsc4A%>" <%if Tsc4A<>0 then response.write " checked" end if%>    onclick="{clearTs4(10)}">銀行轉帳超額細明表</DD>    
                             <DD><input type="checkbox" name="Tsc4B" value="<%=Tsc4B%>" <%if Tsc4B<>0 then response.write " checked" end if%>    onclick="{clearTs4(11)}">庫房脫期建立</DD>
                             
                             <DD><input type="checkbox" name="Tsc4C" value="<%=Tsc4C%>" <%if Tsc4C<>0 then response.write " checked" end if%>    onclick="{clearTs4(12}">庫房過數</DD>
                             <DD><input type="checkbox" name="Tsc4D" value="<%=Tsc4D%>" <%if Tsc4D<>0 then response.write " checked" end if%>    onclick="{clearTs4(13)}">庫房轉帳試算</DD>

                            </Ul></td> 
                 </tr>
                <tr>                                    
                       <td width="300" ><ul><input type="checkbox" name="Tsc8" value="<%=Tsc8%>" <%if Tsc8<>0 then response.write " checked" end if%>    onclick="{clearTs8(0)}">分析及統計

                             <DD><input type="checkbox" name="Tsc81" value="<%=Tsc81%>" <%if Tsc81<>0 then response.write " checked" end if%>    onclick="{clearTs8(1)}">社員統計資料分部報告</DD>
                             <DD><input type="checkbox" name="Tsc82" value="<%=Tsc82%>" <%if Tsc82<>0 then response.write " checked" end if%>    onclick="{clearTs8(2)}">社員報告(保險)</DD>    
                             <DD><input type="checkbox" name="Tsc83" value="<%=Tsc83%>" <%if Tsc83<>0 then response.write " checked" end if%>    onclick="{clearTs8(3)}">社員報告(註冊官)</DD>
                             <DD>&nbsp</DD>
 
                           </Ul></td>
                 </tr>
               <tr>                                    
                       <td width="300" ><ul><input type="checkbox" name="Tsc9" value="<%=Tsc9%>" <%if Tsc9<>0 then response.write " checked" end if%>    onclick="{clearTs9(0);}">系統維護

                             <DD><input type="checkbox" name="Tsc91" value="<%=Tsc91%>" <%if Tsc91<>0 then response.write " checked" end if%>    onclick="{clearTs9(1)}">資料庫輸出</DD>
                             <DD><input type="checkbox" name="Tsc92" value="<%=Tsc92%>" <%if Tsc92<>0 then response.write " checked" end if%>    onclick="{clearTs9(2)}">資料庫輸入</DD>    
                             <DD><input type="checkbox" name="Tsc93" value="<%=Tsc93%>" <%if Tsc93<>0 then response.write " checked" end if%>    onclick="{clearTs9(3)}">用戶管理-新增</DD>
                             <DD><input type="checkbox" name="Tsc94" value="<%=Tsc94%>" <%if Tsc94<>0 then response.write " checked" end if%>    onclick="{clearTs9(4)}">用戶管理-修改</DD>  
                             <DD><input type="checkbox" name="Tsc95" value="<%=Tsc95%>" <%if Tsc95<>0 then response.write " checked" end if%>    onclick="{clearTs9(5)}">用戶管理-更改密碼</DD>
                             <DD><input type="checkbox" name="Tsc96" value="<%=Tsc96%>" <%if Tsc96<>0 then response.write " checked" end if%>    onclick="{clearTs9(6)}">用戶使用紀錄</DD>
                             <DD>&nbsp</DD>
 
                           </Ul></td>
                 </tr>
             </table>
       
      </td>
    </tr>
	<tr>
		<td colspan="9" align="right" valign="middle">
			<%if session("userLevel")<>5 then%>
			<input type="submit" value="儲存" onclick="return validating()&&confirm('確定儲存?')" name="action" class="sbttn">
			<%end if%>
<%if uid="" then %>
		        <input type="button" value="取消" name="bye"  class="sbttn">
<%end if %>
			<input type="button" value="返回" name="back" class="sbttn">
		</td>
	</tr>
</table>
</form>
</center>
</body>
</html>