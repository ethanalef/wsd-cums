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
		<% if Session("userRight_x_Member1") Then %><tr><td colspan='2'><A href='MemberAdd2.asp'>�[�J�s����</A></td></tr><% End If %>
		<% if Session("userRight_x_Member2") Then %><tr><td colspan='2'><A href='MemberMod2.asp'>������ƭץ�</A></td><% End If %>
		<% if Session("userRight_x_Member3") Then %><tr><td colspan='2'><A href='chgroup.asp'>�ഫ�p���H�إ�</A></td></tr><% End If %>
		<% if Session("userRight_x_Member4") Then %><tr><td colspan='2'><A href='bank.asp'>�Ȧ��ƾާ@</A></td></tr><% End If %>
	
		<% if Session("userRight_x_Member5") Then %><tr><td colspan='2'><A href='newacc.asp'>�s�����}��إ�</A></td></tr><% End If %>
	
		<% if Session("userRight_x_Member6") Then %>
		<tr><td colspan='2'><HR STYLE=' color: #CCCCCC' SIZE=1 width='100%'></td></tr>
		
		<tr><td colspan='2'><A href='cutperiod.asp'>�I�Ƴ]�w�إ�</A></td></tr> 
		<% End If %>
	</table>
</DIV>
<DIV id='menu2' class='menu' onMouseover='activateMenu(2);'>
	<table cellpadding='0' cellspacing='0'>
		<% if Session("userRight_x_Loan1") Then %><tr><td colspan='2'><A href='loan.asp'>�U�ڥӽ�</A></td></tr><% End If %>
		<% if Session("userRight_x_Loan2") Then %><tr><td colspan='2'><A href='nloanDetail.asp'>�s�U�ګإ�</A></td></tr><% End If %>
		<% if Session("userRight_x_Loan3") Then %><tr><td colspan='2'><A href='ncloandetail.asp'>�U�ڭץ�</A></td></tr><% End If %>
		<% if Session("userRight_x_Loan4") Then %><tr><td colspan='2'><A href='lnlst.asp'>�U�ڦC�L</A></td></tr><% End If %>
		<% if Session("userRight_x_Loan5") Then %>
		    <tr><td colspan='2'><HR STYLE=' color: #CCCCCC' SIZE=1 width='100%'></td></tr>
		    <tr><td colspan='2'><A href='delayPro.asp'>�����ާ@</A></td></tr><% End If %>
		<% if Session("userRight_x_Loan6") Then %>
		    <tr><td colspan='2'><HR STYLE=' color: #CCCCCC' SIZE=1 width='100%'></td></tr>
		    <tr><td colspan='2'><A href='repayloan.asp'>�{���ٴ�</A></td></tr>
		<% End If %>
		<% if Session("userRight_x_Loan7") Then %><tr><td colspan='2'><A href='saveloan.asp'>�Ѫ��ٴ�</A></td></tr><% End If %>
		<% if Session("userRight_x_Loan8") Then %><tr><td colspan='2'><A href='repbackloan.asp'>�U�ڰh�ڦܪѪ��ާ@</A></td></tr><% End If %>
                <% if Session("userRight_x_Loan9") Then %><tr><td colspan='2'><A href='wofflnb.asp'>�����U�ګإ�</A></td></tr><% End If %>
		<% if Session("userRight_x_Loan10") Then %><tr><td colspan='2'><A href='lntlst.asp'>�U�ڲӶ��C�L</A></td></tr><% End If %>
		<% if Session("userRight_x_Loan11") Then %><tr><td colspan='2'>
		<HR STYLE=' color: #CCCCCC' SIZE=1 width='100%'></td></tr>
		<tr><td colspan='2'><A href='loanadj.asp'>�U�ڲӶ��ץ�</A></td></tr>
		<% End If %>
		<% if Session("userRight_x_Loan12") Then %><tr><td colspan='2'><A href='CanrejA.asp'>�����Ȧ����إ�</A></td></tr><% End If %>
	</table>
</DIV>
<DIV id='menu3' class='menu' onMouseover='activateMenu(3);'>
	<table cellpadding='0' cellspacing='0'>
		<% if Session("userRight_x_cLoan1") Then %><tr><td colspan='2'><A href='lcloan.asp'>�`���U��</A></td></tr><% End If %>
		<% if Session("userRight_x_cLoan2") Then %><tr><td colspan='2'><A href='ccloan.asp'>�{���M��</A></td></tr><% End If %>
		<% if Session("userRight_x_cLoan3") Then %><tr><td colspan='2'><A href='shwdloan.asp'>�Ѫ��M��</A></td></tr><% End If %>
		<% if Session("userRight_x_cLoan4") Then %><tr><td colspan='2'><A href='scloan.asp'>�{���M��(����)</A></td></tr><% End If %>
		<% if Session("userRight_x_cLoan5") Then %><tr><td colspan='2'>
		<HR STYLE=' color: #CCCCCC' SIZE=1 width='100%'></td></tr>		
		<tr><td colspan='2'><A href='crash.asp'>�}���ާ@�إ�</A></td></tr>   
		<% End If %>
                <% if Session("userRight_x_cLoan6") Then %><tr><td colspan='2'><A href=''chlst.asp'>�}���C�L</A></td></tr><% End If %> 
		<% if Session("userRight_x_cLoan7") Then %><tr><td colspan='2'>
		<HR STYLE=' color: #CCCCCC' SIZE=1 width='100%'></td></tr>		
		<tr><td colspan='2'><A href='IvaPro.asp'>IVA�ާ@�إ�</A></td></tr>   
		<% End If %>
                <% if Session("userRight_x_cLoan8") Then %><tr><td colspan='2'><A href='IvaPlst.asp'>IVA�C�L</A></td></tr><% End If %> 
	</table>
</DIV>

<DIV id='menu4' class='menu' onMouseover='activateMenu(4);'>
	<table cellpadding='0' cellspacing='0'>
		<% if Session("userRight_x_AutoPay1") Then %><tr><td colspan='2'><A href='nautopay3.asp'>��b�إ�</A></td></tr><% End If %>
		<% if Session("userRight_x_AutoPay2") Then %><tr><td colspan='2'><A href='Mautopay.asp'>�S�O�Ӯ���b��J�ާ@</A></td></tr><% End If %>
                <% if Session("userRight_x_AutoPay3") Then %><tr><td colspan='2'><A href='atList.asp'>�Ȧ���b�պ�</A></td></tr><% End If %>
		<% if Session("userRight_x_AutoPay4") Then %><tr><td colspan='2'><A href='plnlst.asp'>�S�O�Ӯ���b�պ�</A></td></tr><% End If %>
		<% if Session("userRight_x_AutoPay5") Then %><tr><td colspan='2'><A href='autopass.asp'>�Ȧ���b�ϺЫإ�</A></td></tr><% End If %>
		<% if Session("userRight_x_AutoPay6") Then %><tr><td colspan='2'><A href='AutoAdkt.asp'>�Ȧ����إ�</A></td></tr><% End If %>
		
		<% if Session("userRight_x_AutoPay7") Then %><tr><td colspan='2'><A href='autoupd.asp'>�Ȧ���b�L�� </A></td></tr><% End If %>
                <% if Session("userRight_x_AutoPay8") Then %><tr><td colspan='2'><A href='autolstpro.asp'>�Ȧ�۰���b���ĳq���ѫإ�</A></td></tr><% End If %>
                <% if Session("userRight_x_AutoPay9") Then %><tr><td colspan='2'><A href='RejectLst.asp'>�Ȧ�۰���b���ĳq���ѦC�L</A></td></tr><% End If %>
                <% if Session("userRight_x_AutoPay10") Then %><tr><td colspan='2'><A href='atovList.asp'>�Ȧ���b�W�B�ө���</A></td></tr><% End If %> 
		<% if Session("userRight_x_AutoPay11") Then %><tr><td colspan='2'>
		<tr><td colspan='2'><HR STYLE=' color: #CCCCCC' SIZE=1 width='100%'></td></tr>
		<tr><td colspan='2'><A href='AutoBdkt.asp'>�w�в���إ�</A></td></tr>
		<% End If %>
		
		<% if Session("userRight_x_AutoPay12") Then %><tr><td colspan='2'><A href='sadtupd.asp'>�w�йL��</A></td></tr><% End If %>
                <% if Session("userRight_x_AutoPay13") Then %><tr><td colspan='2'><A href='sdList.asp'>�w����b�պ�</A></td></tr><% End If %>
		
               
                
               
	</table>
</DIV>
<DIV id='menu5' class='menu' onMouseover='activateMenu(5);'>
	<table cellpadding='0' cellspacing='0'>
		<% if Session("userRight_x_Saving1") Then %><tr><td colspan='2'><A href='dvdcal.asp'>�Ѯ��p��ާ@</A></td></tr><% End If %>
		<% if Session("userRight_x_Saving2") Then %><tr><td colspan='2'><A href='divdlist.asp'>�Ѯ��C�L</A></td></tr><% End If %>
		<% if Session("userRight_x_Saving3") Then %><tr><td colspan='2'><A href='Separat.asp'>�������t�إ�</A></td></tr><% End If %>
                <% if Session("userRight_x_Saving3") Then %><tr><td colspan='2'><A href='shpayPro.asp'>�������t�ק�ާ@</A></td></tr><% End If %>
		<% if Session("userRight_x_Saving4") Then %><tr><td colspan='2'><A href='ShAupass.asp'>�Ȧ欣���ϺЫإ�</A></td></tr><% End If %>
		<% if Session("userRight_x_Saving5") Then %><tr><td colspan='2'><A href='divuptd.asp'>�����L��</A></td></tr><% End If %>		
                <% if Session("userRight_x_Saving11") Then %><tr><td colspan='2'><A href='shpreject.asp'>�Ȧ���b���īإ�</A></td></tr><% End If %>
                <% if Session("userRight_x_Saving6") Then %><tr><td colspan='2'><A href='divHuptd.asp'>�Ȱ������L��</A></td></tr><% End If %>		
		<% if Session("userRight_x_Saving7") Then %>
		<tr><td colspan='2'><HR STYLE=' color: #CCCCCC' SIZE=1 width='100%'></td></tr>
		<tr><td colspan='2'><A href='savewithd.asp'>�h�ѫإ�</A></td></tr>
		<% End If %>
		<% if Session("userRight_x_Saving8") Then %><tr><td colspan='2'><A href='savecash.asp'>�{���s�ګإ�</A></td></tr><% End If %>
		<% if Session("userRight_x_Saving9") Then %><tr><td colspan='2'><A href='savtlst.asp'>�Ѫ��C�L</A></td></tr><% End If %>
		<% if Session("userRight_x_Saving10") Then %>
		<tr><td colspan='2'><HR STYLE=' color: #CCCCCC' SIZE=1 width='100%'></td></tr>
		<tr><td colspan='2'><A href='saveadjA.asp'>�Ѫ��Ӷ��ץ�</A></td></tr>
		<% End If %>
	</table>
</DIV>
<DIV id='menu6' class='menu' onMouseover='activateMenu(6);'>
    <table cellpadding='0' cellspacing='0'>
        <% if Session("userRight_x_MemAcct1") Then %><tr><td colspan='2'><A href='acdetail2.asp'>������Ƭd��</A></td></tr><% End If %>
    </table>
</DIV>
<DIV id='menu7' class='menu' onMouseover='activateMenu(7);'>
	<table cellpadding='0' cellspacing='0'>
		<% if Session("userRight_x_Reporting1") Then %><tr><td colspan='2'><A href='acdetaillst.asp'>�ӤH��ƦC��</A></td></tr><% End If %>
		<% if Session("userRight_x_Reporting2") Then %><tr><td colspan='2'><A href='delinquentReport.asp'>�b�b���i</A></td></tr><% End If %>
		<% if Session("userRight_x_Reporting3") Then %><tr><td colspan='2'><A href='dormantList.asp'>�N����i</A></td></tr><% End If %>
		<% if Session("userRight_x_Reporting4") Then %><tr><td colspan='2'><A href='ivalst.asp'>IVA���i</A></td></tr><% End If %>
                <% if Session("userRight_x_Reporting15") Then %><tr><td colspan='2'><A href='carshlst.asp'>�}�����i</A></td></tr> <% End If %> 
		<% if Session("userRight_x_Reporting5") Then %><tr><td colspan='2'><A href='sectionList.asp'>��������/�խ��C��</A></td></tr><% End If %>
		<% if Session("userRight_x_Reporting6") Then %><tr><td colspan='2'><A href='MemDlst.asp'>������b��ƦC��</A></td></tr><% End If %>
		<% if Session("userRight_x_Reporting7") Then %><tr><td colspan='2'><A href='birthdayList.asp'>�����ͤ�W��</A></td></tr><% End If %>
                <% if Session("userRight_x_Reporting8") Then %><tr><td colspan='2'><A href='retirelst.asp'>�h��������i</A></td></tr><% End If %>  
                <% if Session("userRight_x_Reporting9") Then %><tr><td colspan='2'><A href='memstlst.asp'>�������p�C�L</A></td></tr><% End If %>
		<% if Session("userRight_x_Reporting10") Then %>
		<tr><td colspan='2'><HR STYLE=' color: #CCCCCC' SIZE=1 width='100%'></td></tr>
		<tr><td colspan='2'><A href='monCtlst.asp'>�{���b�C��</A></td></tr>
		<% End If %>
		<% if Session("userRight_x_Reporting11") Then %><tr><td colspan='2'><A href='monTtlst.asp'>�w�бb�C��</A></td></tr><% End If %>
		<% if Session("userRight_x_Reporting12") Then %><tr><td colspan='2'><A href='monBtlst.asp'>�Ȧ�b�C��</A></td></tr><% End If %>
		<% if Session("userRight_x_Reporting13") Then %><tr><td colspan='2'><A href='monOtlst.asp'>��L�b�C</A></td></tr><% End If %>
		<% if Session("userRight_x_Reporting14") Then %><tr><td colspan='2'><A href='balList.asp'>�C��b�έp�C��</A></td></tr><% End If %>
                <% if Session("userRight_x_Reporting15") Then %><tr><td colspan='2'><A href='Hyprt.asp'>�b�~��</A></td></tr><% End If %>
	</table>
</DIV>


<DIV id='menu8' class='menu' onMouseover='activateMenu(8);'>
	<table cellpadding='0' cellspacing='0'>
		<% if Session("userRight_x_statist1") Then %><tr><td colspan='2'><A href='InsurLst.asp'>�����έp��Ƥ������i</A></td></tr><% End If %>
                <% if Session("userRight_x_statist2") Then %><tr><td colspan='2'><A href='memIlst.asp'>�������i(�O�I)</A></td></tr><% End If %>
                <% if Session("userRight_x_statist3") Then %><tr><td colspan='2'><A href='memRlst.asp'>�������i(���U�x)</A></td></tr><% End If %> 


	</table>
</DIV>

<DIV id='menu9' class='menu' onMouseover='activateMenu(9);'>
	<table cellpadding='0' cellspacing='0'>
		<% if Session("userRight_x_Other1")Then %><tr><td colspan='2'><A href='dataExport.asp'>��Ʈw��X</A></td></tr><% End If %>
		<% if Session("userRight_x_Other2")Then %><tr><td colspan='2'><A href='dataImport.asp'>��Ʈw��J</A></td></tr><% End If %>
		<% if Session("userRight_x_Other3")Then %><tr><td colspan='2'><A href='userAdd.asp'>�Τ�޲z-�s�W</A></td></tr><% End If %>
		<% if Session("userRight_x_Other3")Then %><tr><td colspan='2'><A href='userMod.asp'>�Τ�޲z-�ק�</A></td></tr><% End If %>
		<% if Session("userRight_x_Other3")Then %><tr><td colspan='2'><A href='chgpass.asp'>���K�X</A></td></tr><% End If %>
		<% if Session("userRight_x_Other4")Then %><tr><td colspan='2'><A href='userLog.asp'>�Τ�ϥά���</A></td></tr><% End If %>
	</table>
</DIV>



<table>
<tr>
<td><img src="images/logo.gif"></td>
<td>���ȸp���u�x�W���U���t��<br>
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

jo_menu_text = "�������,�U��,�M�Ƥί}���ާ@,�۰���b,�Ѫ�,�ӤH��f,����,���R�βέp,�t�κ��@,�n�X "
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
