
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

<%


sub draw_sub_menu()

	jo_sub_menu_array = split(jo_sub_menu_text,",")
	jo_sub_menu_link_array = split(jo_sub_menu_link,",")

	Response.Write	"<DIV id='menu" & did & "' class='menu' onMouseover='activateMenu(" & did & ");'>"
	Response.Write	"<table cellpadding='0' cellspacing='0'>"

	For i = 0 To UBound(jo_sub_menu_array)


		If jo_sub_menu_array(i) = "hr" Then	
	
			Response.Write	"<tr><td colspan='2'><HR STYLE=' color: #CCCCC' SIZE=1 width='100%'></td></tr>"
		Else
			Response.Write	"<tr><td colspan='2'><A href='" & jo_sub_menu_link_array(i) & "'>" & jo_sub_menu_array(i) & "</A></td></tr>"

		End If	

	Next

	Response.Write	"</table>"
	Response.Write	"</DIV>"

end sub




select case session("userlevel")
case 5 ''Auditor

case 4 ''Administrator


	'jo_sub_menu_text = "Member Maintenance,Member list,Birthday list,hr,Delinquent Loan report,Dormant Account List,Section report"
	'jo_sub_menu_link = "member.asp,memList.asp,birthdayList.asp,hr,delinquentReport.asp,dormantList.asp,sectionList.asp"

	jo_sub_menu_text = "�[�J�s����,������ƭץ�,�ഫ�p���H�إ�,hr,�s�����}��إ�"
	jo_sub_menu_link =  "memberAdd2.asp, MemberMod2.asp,chgroup.asp,hr,newacc.asp"

        did = 1

	draw_sub_menu()



	'jo_sub_menu_text = "Loan application Assessment,Loan Application Assessment Reports" 
	jo_sub_menu_text = "�U�ڥӽ�,�s�U�ګإ�,�U�ڭץ�,�U�ڦC�L,hr,�����ާ@,hr,�{���ٴ�,�Ѫ��ٴ�,�U�ڰh�ڦܪѪ��ާ@,�U�ڲӶ��C�L,hr,�U�ڲӶ��ץ�"
	jo_sub_menu_link = "loan.asp,nloanDetail.asp,ncloandetail.asp,lnlst.asp,hr,delayPro.asp,hr,repayloan.asp,saveloan.asp,repbackloan.asp,lntlst.asp,hr,loanadj.asp"

        did = 2

	draw_sub_menu()

	'jo_sub_menu_text = "Loan application Assessment,Loan Application Assessment Reports" 
	jo_sub_menu_text = "�`���U��,�{���M��,�Ѫ��M��"
	jo_sub_menu_link = "lcloan.asp,ccloan.asp,shwdloan.asp"

        did = 3

	draw_sub_menu()	'jo_sub_menu_text = "Loan application Assessment,Loan Application Assessment Reports" 
	jo_sub_menu_text = "��b�إ�,�S�O�Ӯ���b��J�ާ@,�Ȧ���b�պ�,�Ȧ���b�ϺЫإ�,�Ȧ����إ�,�Ȧ���b�W�B�ө���,�Ȧ���b�L�� ,hr,�w�в���إ�,�w����b�պ�,�w�йL��,hr,�S�O�Ӯ���b�պ�"
	jo_sub_menu_link = "nautopay3.asp,Mautopay.asp,atList.asp,Autopass.asp,AutoAdkt.asp,atovList.asp,autoupd.asp,hr,AutoBdkt.asp,sdList.asp,sadtupd.asp,hr,plnlst.asp"

        did = 4

	draw_sub_menu()

	'jo_sub_menu_text = "Loan application Assessment,Loan Application Assessment Reports" 
	jo_sub_menu_text = "�Ѯ��p��ާ@,�Ѯ��C�L,�Ѯ����t�إ�,�Ѯ����t�C�L,�Ѯ��L��,�Ȧ欣���ϺЫإ�,hr,�h�ѫإ�,�{���s�ګإ�,�Ѫ��C�L,hr,�Ѫ��Ӷ��ץ�"
	jo_sub_menu_link = "dvdcal.asp,divdlist.asp,separat.asp,divlst.asp,divpass.asp,Shaupass.asp,hr,savewithd.asp,savecash.asp,savtlst.asp,hr,saveadjA.asp"

        did = 5

	draw_sub_menu()

	'jo_sub_menu_text = "Account Enquiry,Account Maintenance,Daily Transaction Input,Print Balance Statement,Transaction List,hr,League-Due Process,Auto-processing for Auto-pay,Auto-processing for Salary deduction,Account Check list for Auto-pay,Account Check list for Salary deduction,hr,Year End Testing Report,Year End Process,hr,Meeting Notes,Monthly Report,Committee,hr,Cheque Reconciliation"
	'jo_sub_menu_link = "ac.asp,ac.asp,acTx.asp,acBal.asp,acTxList.asp,hr,leagueDue.asp,autoplay.asp,salaryDeduction.asp,atList.asp,sdList.asp,hr,yearEndReport.asp,yearEnd.asp,hr,meetingNotes.asp,monthlyReport.asp,handleParty.asp,hr,cheque.asp"

	jo_sub_menu_text = "������Ƭd��"
	jo_sub_menu_link = "acdetail2.asp"

        did = 6

	draw_sub_menu()

	'jo_sub_menu_text = "Print Balance Statement,A/C Transaction List,Account Check list for Auto-pay,Account Check list for Salary deduction,Delinquent Loan report,Dormant Account List,Section report,Year End Testing Report,Loan Application Assessment Reports,Member list,Birthday list"
	'jo_sub_menu_link = "acBal.asp,acTxList.asp,atList.asp,sdList.asp,delinquentReport.asp,dormantList.asp,sectionList.asp,yearEndReport.asp,loanReport.asp,memList.asp,birthdayList.asp"

	jo_sub_menu_text = "�ӤH��ƦC��,�b����i,�N����i,�p�����C��,������b��ƦC��,�����ͤ�W��,hr,�{���b�C��,�w�бb�C��,�Ȧ�b�C��,��L�b�C,�C��b�έp�C��"
	jo_sub_menu_link = "acdetaillst.asp,delinquentReport.asp,dormantList.asp,sectionList.asp,memDlst.asp,birthdayList.asp,hr,monCtlst.asp,monTtlst.asp,monBtlst.asp,monOtlst.asp,balList.asp"

        did = 7
	draw_sub_menu()




	'jo_sub_menu_text = "G/L Maintenance,G/L List,Transaction Maintenance,Transaction List,Post Surplus or Deficit,Post Delinquent,hr,Period End,hr,Trial Balance,Profit & Lost Statement,Balance Sheet"
	'jo_sub_menu_link = "gl.asp,glList.asp,glTx.asp,glTxList.asp,postSurplus.asp,postDelinquent.asp,hr,periodEnd.asp,hr,trialBalance.asp,pl.asp,balanceSheet.asp"

	'jo_sub_menu_text = "�`�b�ץ�,�`�b��,�`�b�J��,�`�b�Ӷ��C�L,�k�ǬէQ,�k�ǧb��,hr,�C�맹��,hr,�`�b�պ��,�l�q��,�]�ȳ��i��,�����ͤ�W��,�`�b�ץ�,�`�b��,�`�b�J��,�`�b�Ӷ��C�L,�`�b�պ��,�l�q��,�]�ȳ��i��"
	'jo_sub_menu_link = "gl.asp,glList.asp,glTx.asp,glTxList.asp,postSurplus.asp,postDelinquent.asp,hr,periodEnd.asp,hr,trialBalance.asp,pl.asp,balanceSheet.asp,birthdayList.asp.gl.asp,glList.asp,glTx.asp,glTxList.asp,hr,trialBalance.asp,pl.asp,balanceSheet.asp"
         jo_sub_menu_text = ""
         jo_sub_menu_link =""
         
         did = 8

	draw_sub_menu()



	'jo_sub_menu_text = "Database Exporting,Database Importing,User Administration,User Log Report"
	'jo_sub_menu_link = "dataExport.asp,dataImport.asp,user.asp,userLog.asp"

	jo_sub_menu_text = "��Ʈw��X,��Ʈw��J,�Τ�޲z,�Τ�ϥά���"
	jo_sub_menu_link = "dataExport.asp,dataImport.asp,user.asp,userLog.asp"


	did = 9

	draw_sub_menu()


case 3 ''Supervisor

	'jo_sub_menu_text = "Member Maintenance,Member list,Birthday list,hr,Delinquent Loan report,Dormant Account List,Section report"
	'jo_sub_menu_link = "member.asp,memList.asp,birthdayList.asp,hr,delinquentReport.asp,dormantList.asp,sectionList.asp"

	jo_sub_menu_text = "�[�J�s����,������ƭץ�,�ഫ�p���H�إ�,hr,�s�����}��إ�"
	jo_sub_menu_link =  "memberAdd2.asp, MemberMod2.asp,chgroup.asp,hr,newacc.asp"

        did = 1

	draw_sub_menu()



	'jo_sub_menu_text = "Loan application Assessment,Loan Application Assessment Reports" 
	jo_sub_menu_text = "�U�ڥӽ�,�s�U�ګإ�,�U�ڭץ�,�U�ڦC�L,hr,�����ާ@,hr,�{���ٴ�,�Ѫ��ٴ�,�U�ڰh�ڦܪѪ��ާ@,�U�ڲӶ��C�L"
	jo_sub_menu_link = "loan.asp,nloanDetail.asp,ncloandetail.asp,lnlst.asp,hr,delayPro.asp,hr,repayloan.asp,saveloan.asp,repbackloan.asp,lntlst.asp"

        did = 2

	draw_sub_menu()

	'jo_sub_menu_text = "Loan application Assessment,Loan Application Assessment Reports" 
	jo_sub_menu_text = "�`���U��,�{���M��,�Ѫ��M��"
	jo_sub_menu_link = "lcloan.asp,ccloan.asp,shwdloan.asp"

        did = 3

	draw_sub_menu()	'jo_sub_menu_text = "Loan application Assessment,Loan Application Assessment Reports" 
	jo_sub_menu_text = "��b�إ�,�S�O�Ӯ���b��J�ާ@,�Ȧ���b�պ�,�Ȧ���b�ϺЫإ�,�Ȧ����إ�,�Ȧ���b�W�B�ө���,�Ȧ���b�L�� ,hr,�w�в���إ�,�w����b�պ�,�w�йL��,hr,�S�O�Ӯ���b�պ�"
	jo_sub_menu_link = "nautopay3.asp,Mautopay.asp,atList.asp,Autopass.asp,AutoAdkt.asp,atovList.asp,autoupd.asp,hr,AutoBdkt.asp,sdList.asp,sadtupd.asp,hr,plnlst.asp"

        did = 4

	draw_sub_menu()

	'jo_sub_menu_text = "Loan application Assessment,Loan Application Assessment Reports" 
	jo_sub_menu_text = "�Ѯ��p��ާ@,�Ѯ��C�L,�Ѯ����t�إ�,�Ѯ����t�C�L,�Ѯ��L��,�Ȧ欣���ϺЫإ�,hr,�h�ѫإ�,�{���s�ګإ�,�Ѫ��C�L"
	jo_sub_menu_link = "dvdcal.asp,divdlist.asp,separat.asp,divlst.asp,divpass.asp,Shaupass.asp,hr,savewithd.asp,savecash.asp,savtlst.asp"

        did = 5

	draw_sub_menu()

	'jo_sub_menu_text = "Account Enquiry,Account Maintenance,Daily Transaction Input,Print Balance Statement,Transaction List,hr,League-Due Process,Auto-processing for Auto-pay,Auto-processing for Salary deduction,Account Check list for Auto-pay,Account Check list for Salary deduction,hr,Year End Testing Report,Year End Process,hr,Meeting Notes,Monthly Report,Committee,hr,Cheque Reconciliation"
	'jo_sub_menu_link = "ac.asp,ac.asp,acTx.asp,acBal.asp,acTxList.asp,hr,leagueDue.asp,autoplay.asp,salaryDeduction.asp,atList.asp,sdList.asp,hr,yearEndReport.asp,yearEnd.asp,hr,meetingNotes.asp,monthlyReport.asp,handleParty.asp,hr,cheque.asp"

	jo_sub_menu_text = "������Ƭd��"
	jo_sub_menu_link = "acdetail2.asp"

        did = 6

	draw_sub_menu()

	'jo_sub_menu_text = "Print Balance Statement,A/C Transaction List,Account Check list for Auto-pay,Account Check list for Salary deduction,Delinquent Loan report,Dormant Account List,Section report,Year End Testing Report,Loan Application Assessment Reports,Member list,Birthday list"
	'jo_sub_menu_link = "acBal.asp,acTxList.asp,atList.asp,sdList.asp,delinquentReport.asp,dormantList.asp,sectionList.asp,yearEndReport.asp,loanReport.asp,memList.asp,birthdayList.asp"

	jo_sub_menu_text = "�ӤH��ƦC��,�b����i,�N����i,�p�����C��,������b��ƦC��,�����ͤ�W��,hr,�{���b�C��,�w�бb�C��,�Ȧ�b�C��,��L�b�C,�C��b�έp�C��"
	jo_sub_menu_link = "acdetaillst.asp,delinquentReport.asp,dormantList.asp,sectionList.asp,memDlst.asp,birthdayList.asp,hr,monCtlst.asp,monTtlst.asp,monBtlst.asp,monOtlst.asp,balList.asp"

        did = 7
	draw_sub_menu()







case 2 ''Operator

	jo_sub_menu_text = "�[�J�s����,������ƭץ�,�ഫ�p���H�إ�,hr,�s�����}��إ�"
	jo_sub_menu_link =  "memberAdd2.asp, MemberMod2.asp,chgroup.asp,hr,newacc.asp"

        did = 1

	draw_sub_menu()

	jo_sub_menu_text = "������Ƭd��"
	jo_sub_menu_link = "acdetail2.asp"

        did = 2

	draw_sub_menu()

end select
%>





<%


select case session("userlevel")

	case 4 ''Administrator

	'jo_menu_text  = "Membership,Loan,Account System,General Ledger,Reports shortcut,Others"
	jo_menu_text = "�������,�U��,�M�ƾާ@,�۰���b,�Ѫ�,�ӤH��f,����,�`�b,��L,�n�X "
	jo_menu_array = split(jo_menu_text, ",")

        case 3  ''supervisor

	'jo_menu_text  = "Membership,Loan,Account System,General Ledger,Reports shortcut,Others"
	jo_menu_text = "�������,�U��,�M�ƾާ@,�۰���b,�Ѫ�,�ӤH��f,����,�`�b,��L,�n�X "
	jo_menu_array = split(jo_menu_text, ",")

        case 2  ''Operator
	'jo_menu_text  = "Membership,Loan,Account System,General Ledger,Reports shortcut,Others"
	jo_menu_text = "�������,�ӤH��f,�n�X "  
        jo_menu_array = split(jo_menu_text, ",")

end select

%>




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