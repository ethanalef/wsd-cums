<link rel="stylesheet" href="template1.css" type="text/css">
<link rel="stylesheet" href="main.css" type="text/css">

<SCRIPT language="javascript">
<!--
var activeMenu = 0;
document.onmouseover = killMenu;
document.onmouseout = killMenu;


function activateMenu(menuLayerRef) {
  if (activeMenu != menuLayerRef) {
    if (activeMenu) hideMenu("menu" + activeMenu);
      activeMenu = menuLayerRef;
      menuID = "menu" + menuLayerRef;
      menutopID = "menutop" + menuLayerRef;
      document.all[menuID].style.pixelTop =  100; //53; //107;
      document.all[menuID].style.pixelLeft = document.all[menutopID].offsetLeft-4;
      showMenu(menuID)
  }
  window.event.cancelBubble = true;
}

function showMenu(layerID) {
  eval('document.all["' + layerID + '"].style.visibility = "visible"');
  ToolbarMenu = document.all(layerID)
  hideElement("SELECT");
  hideElement("OBJECT");
}

function hideMenu(layerID) {
  eval('document.all["' + layerID + '"].style.visibility = "hidden"');
  showElement("SELECT");
  showElement("OBJECT");
}

function killMenu(e) {
  if (activeMenu) {
    menuID = "menu" + activeMenu;
    hideMenu(menuID);
    activeMenu = 0;
  }
}

function hideElement(elmID)
{
  for (i = 0; i < document.all.tags(elmID).length; i++)
  {
    obj = document.all.tags(elmID)[i];
    if (! obj || ! obj.offsetParent)
      continue;

      // Find the element's offsetTop and offsetLeft relative to the BODY tag.
      objLeft   = obj.offsetLeft;
      objTop    = obj.offsetTop;
      objParent = obj.offsetParent;
      while (objParent.tagName.toUpperCase() != "BODY")
      {
        objLeft  += objParent.offsetLeft;
        objTop   += objParent.offsetTop;
        objParent = objParent.offsetParent;
      }
      // Adjust the element's offsetTop relative to the dropdown menu
      objTop = objTop - 69;

      if (ToolbarMenu.offsetLeft > (objLeft + obj.offsetWidth) || objLeft > (ToolbarMenu.offsetLeft + ToolbarMenu.offsetWidth));
      else if (objTop > ToolbarMenu.offsetHeight);
      else
        obj.style.visibility = "hidden";
  	}
}

function showElement(elmID)
{
  for (i = 0; i < document.all.tags(elmID).length; i++)
  {
    obj = document.all.tags(elmID)[i];
    if (! obj || ! obj.offsetParent)
      continue;
    obj.style.visibility = "";
  }
}

<%
thisURL=request.servervariables("script_name")
thisURL=mid(thisURL,InstrRev(thisURL,"/")+1,InstrRev(thisURL,".")-InstrRev(thisURL,"/")-1)
%>
function helppopup(){
  window.open ('../help/<%=thisURL%>.htm','pop','width=700,height=550,statusbar=no,toolbar=no,resizable,scrollbars,dependent')
}
//-->
</script>

<DIV id='menu1' class='menu' onMouseover='activateMenu(1);'>
	<table cellpadding='0' cellspacing='0'>
		<% if Session("userRight_x_Member1") Then %><tr><td colspan='2'><A href='MemberAdd2.asp'>加入新社員</A></td></tr><% End If %>
		<% if Session("userRight_x_Member2") Then %><tr><td colspan='2'><A href='MemberMod2.asp'>社員資料修正</A></td><% End If %>
		<% if Session("userRight_x_Member3") Then %><tr><td colspan='2'><A href='chgroup.asp'>轉換聯絡人建立</A></td></tr><% End If %>
		<% if Session("userRight_x_Member4") Then %><tr><td colspan='2'><A href='bank.asp'>銀行資料操作</A></td></tr><% End If %>
	
		<% if Session("userRight_x_Member5") Then %><tr><td colspan='2'><A href='newacc.asp'>新社員開戶建立</A></td></tr><% End If %>
	
		<% if Session("userRight_x_Member6") Then %>
		<tr><td colspan='2'><HR STYLE=' color: #CCCCCC' SIZE=1 width='100%'></td></tr>
		
		<tr><td colspan='2'><A href='cutperiod.asp'>截數設定建立</A></td></tr> 
		<% End If %>
	</table>
</DIV>
<DIV id='menu2' class='menu' onMouseover='activateMenu(2);'>
	<table cellpadding='0' cellspacing='0'>
		<% if Session("userRight_x_Loan1") Then %><tr><td colspan='2'><A href='loan.asp'>貸款申請</A></td></tr><% End If %>
		<% if Session("userRight_x_Loan2") Then %><tr><td colspan='2'><A href='nloanDetail.asp'>新貸款建立</A></td></tr><% End If %>
		<% if Session("userRight_x_Loan3") Then %><tr><td colspan='2'><A href='ncloandetail.asp'>貸款修正</A></td></tr><% End If %>
		<% if Session("userRight_x_Loan4") Then %><tr><td colspan='2'><A href='lnlst.asp'>貸款列印</A></td></tr><% End If %>
		<% if Session("userRight_x_Loan5") Then %>
		    <tr><td colspan='2'><HR STYLE=' color: #CCCCCC' SIZE=1 width='100%'></td></tr>
		    <tr><td colspan='2'><A href='delayPro.asp'>延期操作</A></td></tr><% End If %>
		<% if Session("userRight_x_Loan6") Then %>
		    <tr><td colspan='2'><HR STYLE=' color: #CCCCCC' SIZE=1 width='100%'></td></tr>
		    <tr><td colspan='2'><A href='repayloan.asp'>現金還款</A></td></tr>
		<% End If %>
		<% if Session("userRight_x_Loan7") Then %><tr><td colspan='2'><A href='saveloan.asp'>股金還款</A></td></tr><% End If %>
		<% if Session("userRight_x_Loan8") Then %><tr><td colspan='2'><A href='repbackloan.asp'>貸款退款至股金操作</A></td></tr><% End If %>
                <% if Session("userRight_x_Loan9") Then %><tr><td colspan='2'><A href='wofflnb.asp'>劃消貸款建立</A></td></tr><% End If %>
		<% if Session("userRight_x_Loan10") Then %><tr><td colspan='2'><A href='lntlst.asp'>貸款細項列印</A></td></tr><% End If %>
		<% if Session("userRight_x_Loan11") Then %><tr><td colspan='2'>
		<HR STYLE=' color: #CCCCCC' SIZE=1 width='100%'></td></tr>
		<tr><td colspan='2'><A href='loanadj.asp'>貸款細項修正</A></td></tr>
		<% End If %>
		<% if Session("userRight_x_Loan12") Then %><tr><td colspan='2'><A href='CanrejA.asp'>取消銀行脫期建立</A></td></tr><% End If %>
	</table>
</DIV>
<DIV id='menu3' class='menu' onMouseover='activateMenu(3);'>
	<table cellpadding='0' cellspacing='0'>
		<% if Session("userRight_x_cLoan1") Then %><tr><td colspan='2'><A href='lcloan.asp'>循環貸款</A></td></tr><% End If %>
		<% if Session("userRight_x_cLoan2") Then %><tr><td colspan='2'><A href='ccloan.asp'>現金清數</A></td></tr><% End If %>
		<% if Session("userRight_x_cLoan3") Then %><tr><td colspan='2'><A href='shwdloan.asp'>股金清數</A></td></tr><% End If %>
		<% if Session("userRight_x_cLoan4") Then %><tr><td colspan='2'><A href='scloan.asp'>現金清數(本金)</A></td></tr><% End If %>
		<% if Session("userRight_x_cLoan5") Then %><tr><td colspan='2'>
		<HR STYLE=' color: #CCCCCC' SIZE=1 width='100%'></td></tr>		
		<tr><td colspan='2'><A href='crash.asp'>破產操作建立</A></td></tr>   
		<% End If %>
                <% if Session("userRight_x_cLoan6") Then %><tr><td colspan='2'><A href=''chlst.asp'>破產列印</A></td></tr><% End If %> 
		<% if Session("userRight_x_cLoan7") Then %><tr><td colspan='2'>
		<HR STYLE=' color: #CCCCCC' SIZE=1 width='100%'></td></tr>		
		<tr><td colspan='2'><A href='IvaPro.asp'>IVA操作建立</A></td></tr>   
		<% End If %>
                <% if Session("userRight_x_cLoan8") Then %><tr><td colspan='2'><A href='IvaPlst.asp'>IVA列印</A></td></tr><% End If %> 
	</table>
</DIV>

<DIV id='menu4' class='menu' onMouseover='activateMenu(4);'>
	<table cellpadding='0' cellspacing='0'>
		<% if Session("userRight_x_AutoPay1") Then %><tr><td colspan='2'><A href='nautopay3.asp'>轉帳建立</A></td></tr><% End If %>
		<% if Session("userRight_x_AutoPay2") Then %><tr><td colspan='2'><A href='Mautopay.asp'>特別個案轉帳輸入操作</A></td></tr><% End If %>
                <% if Session("userRight_x_AutoPay3") Then %><tr><td colspan='2'><A href='atList.asp'>銀行轉帳試算</A></td></tr><% End If %>
		<% if Session("userRight_x_AutoPay4") Then %><tr><td colspan='2'><A href='plnlst.asp'>特別個案轉帳試算</A></td></tr><% End If %>
		<% if Session("userRight_x_AutoPay5") Then %><tr><td colspan='2'><A href='autopass.asp'>銀行轉帳磁碟建立</A></td></tr><% End If %>
		<% if Session("userRight_x_AutoPay6") Then %><tr><td colspan='2'><A href='AutoAdkt.asp'>銀行脫期建立</A></td></tr><% End If %>
		
		<% if Session("userRight_x_AutoPay7") Then %><tr><td colspan='2'><A href='autoupd.asp'>銀行轉帳過數 </A></td></tr><% End If %>
                <% if Session("userRight_x_AutoPay8") Then %><tr><td colspan='2'><A href='autolstpro.asp'>銀行自動轉帳失效通知書建立</A></td></tr><% End If %>
                <% if Session("userRight_x_AutoPay9") Then %><tr><td colspan='2'><A href='RejectLst.asp'>銀行自動轉帳失效通知書列印</A></td></tr><% End If %>
                <% if Session("userRight_x_AutoPay10") Then %><tr><td colspan='2'><A href='atovList.asp'>銀行轉帳超額細明表</A></td></tr><% End If %> 
		<% if Session("userRight_x_AutoPay11") Then %><tr><td colspan='2'>
		<tr><td colspan='2'><HR STYLE=' color: #CCCCCC' SIZE=1 width='100%'></td></tr>
		<tr><td colspan='2'><A href='AutoBdkt.asp'>庫房脫期建立</A></td></tr>
		<% End If %>
		
		<% if Session("userRight_x_AutoPay12") Then %><tr><td colspan='2'><A href='sadtupd.asp'>庫房過數</A></td></tr><% End If %>
                <% if Session("userRight_x_AutoPay13") Then %><tr><td colspan='2'><A href='sdList.asp'>庫房轉帳試算</A></td></tr><% End If %>
		
               
                
               
	</table>
</DIV>
<DIV id='menu5' class='menu' onMouseover='activateMenu(5);'>
	<table cellpadding='0' cellspacing='0'>
		<% if Session("userRight_x_Saving1") Then %><tr><td colspan='2'><A href='dvdcal.asp'>股息計算操作</A></td></tr><% End If %>
		<% if Session("userRight_x_Saving2") Then %><tr><td colspan='2'><A href='divdlist.asp'>股息列印</A></td></tr><% End If %>
		<% if Session("userRight_x_Saving3") Then %><tr><td colspan='2'><A href='Separat.asp'>派息分配建立</A></td></tr><% End If %>
                <% if Session("userRight_x_Saving3") Then %><tr><td colspan='2'><A href='shpayPro.asp'>派息分配修改操作</A></td></tr><% End If %>
		<% if Session("userRight_x_Saving4") Then %><tr><td colspan='2'><A href='ShAupass.asp'>銀行派息磁碟建立</A></td></tr><% End If %>
		<% if Session("userRight_x_Saving5") Then %><tr><td colspan='2'><A href='divuptd.asp'>派息過數</A></td></tr><% End If %>		
                <% if Session("userRight_x_Saving11") Then %><tr><td colspan='2'><A href='shpreject.asp'>銀行轉帳失效建立</A></td></tr><% End If %>
                <% if Session("userRight_x_Saving6") Then %><tr><td colspan='2'><A href='divHuptd.asp'>暫停派息過數</A></td></tr><% End If %>		
		<% if Session("userRight_x_Saving7") Then %>
		<tr><td colspan='2'><HR STYLE=' color: #CCCCCC' SIZE=1 width='100%'></td></tr>
		<tr><td colspan='2'><A href='savewithd.asp'>退股建立</A></td></tr>
		<% End If %>
		<% if Session("userRight_x_Saving8") Then %><tr><td colspan='2'><A href='savecash.asp'>現金存款建立</A></td></tr><% End If %>
		<% if Session("userRight_x_Saving9") Then %><tr><td colspan='2'><A href='savtlst.asp'>股金列印</A></td></tr><% End If %>
		<% if Session("userRight_x_Saving10") Then %>
		<tr><td colspan='2'><HR STYLE=' color: #CCCCCC' SIZE=1 width='100%'></td></tr>
		<tr><td colspan='2'><A href='saveadjA.asp'>股金細項修正</A></td></tr>
		<% End If %>
	</table>
</DIV>
<DIV id='menu6' class='menu' onMouseover='activateMenu(6);'>
    <table cellpadding='0' cellspacing='0'>
        <% if Session("userRight_x_MemAcct1") Then %><tr><td colspan='2'><A href='acdetail2.asp'>社員資料查詢</A></td></tr><% End If %>
    </table>
</DIV>
<DIV id='menu7' class='menu' onMouseover='activateMenu(7);'>
	<table cellpadding='0' cellspacing='0'>
		<% if Session("userRight_x_Reporting1") Then %><tr><td colspan='2'><A href='acdetaillst.asp'>個人資料列表</A></td></tr><% End If %>
		<% if Session("userRight_x_Reporting2") Then %><tr><td colspan='2'><A href='delinquentReport.asp'>呆帳報告</A></td></tr><% End If %>
		<% if Session("userRight_x_Reporting3") Then %><tr><td colspan='2'><A href='dormantList.asp'>冷戶報告</A></td></tr><% End If %>
		<% if Session("userRight_x_Reporting4") Then %><tr><td colspan='2'><A href='ivalst.asp'>IVA報告</A></td></tr><% End If %>
                <% if Session("userRight_x_Reporting15") Then %><tr><td colspan='2'><A href='carshlst.asp'>破產報告</A></td></tr> <% End If %> 
		<% if Session("userRight_x_Reporting5") Then %><tr><td colspan='2'><A href='sectionList.asp'>社員分組/組員列表</A></td></tr><% End If %>
		<% if Session("userRight_x_Reporting6") Then %><tr><td colspan='2'><A href='MemDlst.asp'>社員轉帳資料列表</A></td></tr><% End If %>
		<% if Session("userRight_x_Reporting7") Then %><tr><td colspan='2'><A href='birthdayList.asp'>社員生日名單</A></td></tr><% End If %>
                <% if Session("userRight_x_Reporting8") Then %><tr><td colspan='2'><A href='retirelst.asp'>退休社員報告</A></td></tr><% End If %>  
                <% if Session("userRight_x_Reporting9") Then %><tr><td colspan='2'><A href='memstlst.asp'>社員狀況列印</A></td></tr><% End If %>
		<% if Session("userRight_x_Reporting10") Then %>
		<tr><td colspan='2'><HR STYLE=' color: #CCCCCC' SIZE=1 width='100%'></td></tr>
		<tr><td colspan='2'><A href='monCtlst.asp'>現金帳列表</A></td></tr>
		<% End If %>
		<% if Session("userRight_x_Reporting11") Then %><tr><td colspan='2'><A href='monTtlst.asp'>庫房帳列表</A></td></tr><% End If %>
		<% if Session("userRight_x_Reporting12") Then %><tr><td colspan='2'><A href='monBtlst.asp'>銀行帳列表</A></td></tr><% End If %>
		<% if Session("userRight_x_Reporting13") Then %><tr><td colspan='2'><A href='monOtlst.asp'>其他帳列</A></td></tr><% End If %>
		<% if Session("userRight_x_Reporting14") Then %><tr><td colspan='2'><A href='balList.asp'>每月帳統計列表</A></td></tr><% End If %>
                <% if Session("userRight_x_Reporting15") Then %><tr><td colspan='2'><A href='Hyprt.asp'>半年結</A></td></tr><% End If %>
	</table>
</DIV>


<DIV id='menu8' class='menu' onMouseover='activateMenu(8);'>
	<table cellpadding='0' cellspacing='0'>
		<% if Session("userRight_x_statist1") Then %><tr><td colspan='2'><A href='InsurLst.asp'>社員統計資料分部報告</A></td></tr><% End If %>
                <% if Session("userRight_x_statist2") Then %><tr><td colspan='2'><A href='memIlst.asp'>社員報告(保險)</A></td></tr><% End If %>
                <% if Session("userRight_x_statist3") Then %><tr><td colspan='2'><A href='memRlst.asp'>社員報告(註冊官)</A></td></tr><% End If %> 


	</table>
</DIV>

<DIV id='menu9' class='menu' onMouseover='activateMenu(9);'>
	<table cellpadding='0' cellspacing='0'>
		<% if Session("userRight_x_Other1")Then %><tr><td colspan='2'><A href='dataExport.asp'>資料庫輸出</A></td></tr><% End If %>
		<% if Session("userRight_x_Other2")Then %><tr><td colspan='2'><A href='dataImport.asp'>資料庫輸入</A></td></tr><% End If %>
		<% if Session("userRight_x_Other3")Then %><tr><td colspan='2'><A href='userAdd.asp'>用戶管理-新增</A></td></tr><% End If %>
		<% if Session("userRight_x_Other3")Then %><tr><td colspan='2'><A href='userMod.asp'>用戶管理-修改</A></td></tr><% End If %>
		<% if Session("userRight_x_Other3")Then %><tr><td colspan='2'><A href='chgpass.asp'>更改密碼</A></td></tr><% End If %>
		<% if Session("userRight_x_Other4")Then %><tr><td colspan='2'><A href='userLog.asp'>用戶使用紀錄</A></td></tr><% End If %>
	</table>
</DIV>



<table>
<tr>
<td><img src="images/logo.gif"></td>
<td>水務署員工儲蓄互助社系統<br>
Water Supplies Department Staff Credit Union Membership, Accounting, Savings and Loans Software </td>
</tr>
</table>


<!-- Main Menu -->
<div id="topmenutabs">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
<tr>
<td><img border="0" src="images/blank.gif" /></td>


<!--th class="mtab-ls"><img border="0" src="images/blank.gif" /></th-->
<!--th class="mtab-rs"><img border="0" src="images/blank.gif" /></th-->

<%

jo_menu_text = "社員資料,貸款,清數及破產操作,自動轉帳,股金,個人戶口,報表,分析及統計,系統維護,登出 "
jo_menu_array = split(jo_menu_text, ",")

For i = 0 To (UBound(jo_menu_array)) 
	Response.Write 	"<td class='mtab-l'><img border='0' src='images/blank.gif' /></td>"
  	Response.Write	"<td class='mtab-r'><img border='0' src='images/blank.gif' /></td>"
Next

%>


<td style="width:100%; background: #FFF ;"><img border="0" src="images/blank.gif" /></td>
</tr>
<tr>
<td class="menuBackground">&nbsp;</td>
<%

For i = 0 To (UBound(jo_menu_array))

	if (i = UBound(jo_menu_array)) then
		Response.Write	"<td id='menutop" & (i+1) & "' class='menutabs-td' colspan='2' nowrap>&nbsp;&nbsp;<a href='../logout.asp'>" & jo_menu_array(i) & "</a>&nbsp;&nbsp;</td>"
	else 
		Response.Write	"<td id='menutop" & (i+1) & "' class='menutabs-td' colspan='2' nowrap>&nbsp;&nbsp;<a href='#' onmouseover='activateMenu(" & (i+1) & ");' id='amenutop" & (i+1) & "'>" & jo_menu_array(i) & "</a>&nbsp;&nbsp;</td>"
	end if
Next

%>

<td width="100%" class="menuBackground"><img border="0" src="images/blank.gif" /></td>
</tr>
</table>
</div>
