<%@ Language = "VBScript" %>
<% Response.Buffer = True %>

<html>

<%

' Prepare variables.

Dim oFS, oFSPath
Dim sServername, sServerinst, sPhyspath, sServerVersion 
Dim sServerIP, sRemoteIP
Dim sPath, oDefSite, sDefDoc, sDocName, aDefDoc

Dim bSuccess           ' This value is used later to warn the user if a default document does not exist.
Dim iVer               ' This value is used to pass the server version number to a function.

bSuccess = False
iVer = 0

' Get some server variables to help with the next task.

sServername = LCase(Request.ServerVariables("SERVER_NAME"))
sServerinst = Request.ServerVariables("INSTANCE_ID")
sPhyspath = LCase(Request.ServerVariables("APPL_PHYSICAL_PATH"))
sServerVersion = LCase(Request.ServerVariables("SERVER_SOFTWARE"))
sServerIP = LCase(Request.ServerVariables("LOCAL_ADDR"))      ' Server's IP address
sRemoteIP =  LCase(Request.ServerVariables("REMOTE_ADDR"))    ' Client's IP address

' If the querystring variable uc <> 1, and the user is browsing from the server machine, 
' go ahead and show them localstart.asp.  We don't want localstart.asp shown to outside users.

If Not (sServername = "localhost" Or sServerIP = sRemoteIP) Then
  Response.Redirect "iisstart.asp"
Else 

' Using ADSI, get the list of default documents for this Web site.

sPath = "IIS://" & sServername & "/W3SVC/" & sServerinst
Set oDefSite = GetObject(sPath)
sDefDoc = LCase(oDefSite.DefaultDoc)
aDefDocs = split(sDefDoc, ",")

' Make sure at least one of them is valid.

Set oFS = CreateObject("Scripting.FileSystemObject")

For Each sDocName in aDefDocs
  If oFS.FileExists(sPhyspath & sDocName) Then
    If InStr(sDocName,"iisstart") = 0 Then
      ' IISstart doesn't count because it is an IIS file.
      bSuccess = True  ' This value will be used later to warn the user if a default document does not exist.
      Exit For
    End If
  End If
Next

' Find out what version of IIS is running.

Select Case sServerVersion 
   Case "microsoft-iis/5.0"
     iVer = 50         ' This value is used to pass the server version number to a function.
   Case "microsoft-iis/5.1"
     iVer = 51
   Case "microsoft-iis/6.0"
     iVer = 60
End Select

%>

<head>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">

<script language="javascript">

  // This code is executed before the rest of the page, even before the ASP code above.
  
  var gWinheight;
  var gDialogsize;
  var ghelpwin;
  
  // Move the current window to the top left corner.
  
  window.moveTo(5,5);
  
  // Change the size of the window.

  gWinheight= 480;
  gDialogsize= "width=640,height=480,left=300,top=50,"
  
  if (window.screen.height > 600)
  {
<% if not success and Err = 0 then %>
    gWinheight= 700;
<% else %>
    gWinheight= 700;
<% end if %>
    gDialogsize= "width=640,height=480,left=500,top=50"
  }
  
  window.resizeTo(620,gWinheight);
  
  // Launch IIS Help in another browser window.
  
  loadHelpFront();

function loadHelpFront()
// This function opens IIS Help in another browser window.
{
  ghelpwin = window.open("http://localhost/iishelp/","Help","status=yes,toolbar=yes,scrollbars=yes,menubar=yes,location=yes,resizable=yes,"+gDialogsize,true);  
      window.resizeTo(620,gWinheight);
}

function activate(ServerVersion)
// This function brings up a little help window showing how to open the IIS snap-in.
{
  if (50 == ServerVersion)
    window.open("http://localhost/iishelp/iis/htm/core/iisnapin.htm", "SnapIn", 'toolbar=no, left=200, top=200, scrollbars=yes, resizeable=yes,  width=350, height=350');
  if (51 == ServerVersion)
    window.open("http://localhost/iishelp/iis/htm/core/iiabuti.htm", "SnapIn", 'toolbar=no, left=200, top=200, scrollbars=yes, resizeable=yes,  width=350, height=350');
  if (60 == ServerVersion)
    window.open("http://localhost/iishelp/iis/htm/core/gs_iissnapin.htm", "SnapIn", 'toolbar=no, left=200, top=200, scrollbars=yes, resizeable=yes,  width=350, height=350');
  if (0 == ServerVersion)
    window.open("http://localhost/iishelp/", "Help", 'toolbar=no, left=200, top=200, scrollbars=yes, resizeable=yes,  width=350, height=350');  
}

</script>

<title>�w��ϥ� Windows XP Server Internet �A��</title>
<style>
  ul{margin-left: 15px;}
  .clsHeading {font-family: �s�ө���; color: black; font-size: 11; font-weight: 800; width:210;}  
  .clsEntryText {font-family: �s�ө���; color: black; font-size: 11; font-weight: 400; background-color:#FFFFFF;}    
  .clsWarningText {font-family: �s�ө���; color: #B80A2D; font-size: 11; font-weight: 600; width:550;  background-color:#EFE7EA;}  
  .clsCopy {font-family: �s�ө���; color: black; font-size: 11; font-weight: 400;  background-color:#FFFFFF;}  
</style>
</head>

<body topmargin="3" leftmargin="3" marginheight="0" marginwidth="0" bgcolor="#FFFFFF"
link="#000066" vlink="#000000" alink="#0000FF" text="#000000">

<!-- BEGIN MAIN DOCUMENT BODY --->

<p align="center"><img src="winXP.gif" vspace="0" hspace="0"></p>
<table width="500" cellpadding="5" cellspacing="3" border="0" align="center">

  <tr>
  <td class="clsWarningText" colspan="2">
  
  <table><tr><td>
  <img src="warning.gif" width="40" height="40" border="0" align="left">
  </td><td class="clsWarningText">
  <b>Web �A�Ȳ{�b���椤�C
  
<% If Not bSuccess And Err = 0 Then %>
  
  <p>�z�ثe�èS�����ϥΪ̫إߤ@�ӹw�] Web �����C
������ձq�t�@�x�q���s���z�������ϥΪ̥ثe�|����@�����
  <a href="iisstart.asp?uc=1">�غc��</a> �������C
  �z�� Web ���A���C�X�U�C�ɮ׬��i�઺�w�] Web ����: <%=sDefDoc%>�C�ثe�u�� iisstart.asp �s�b�C<br><br>
  
<% End If %>

  �Y�n�s�W����w�]�����A�бN�ɮ��x�s�b <%=sPhyspath%>�C
  </b>
  </td></tr></table>
 
  </td>
  </tr>
  
  <tr>
  <td>
  <table cellpadding="3" cellspacing="3" border=0 >
  <tr>
    <td valign="top" rowspan=3>
      <img src="web.gif">
    </td>  
    <td valign="top" rowspan=3>
  <span class="clsHeading">
  �w��ϥ� IIS 5.1</span><br>
      <span class="clsEntryText">    
    Internet Information Services (IIS) 5.1 for Microsoft Windows XP Professional
    �N�����B�⪺�¤O 
    �a��F Windows�C���F IIS�A�z�i�H�����a�@���ɮפΦL�����A�Ϊ̱z�i�H�إ����ε{���A
    �b Web �W�w���a�o���T�A�H�W�i�z��´�@�θ�T���覡�CIIS �O�@�Ӧw�������x�A
    �A�X�Ψӫإߤνհt�q�l�ӰȸѨM��סA�H�έ��n�� Web ���ε{���C
  <p>
    �ϥΤw�w�� IIS �� Windows XP Professional�A���Ѥ@�ӭӤH�ζ}�o���@�~�t�ΡA���z�i�H:</span>
  <p>
    <ul class="clsEntryText">
      <li>�w�˭ӤH�� Web ���A��
      <li>�b�p�դ��@�θ�T
      <li>�s����Ʈw
      <li>�}�o���~��������
      <li>�}�o Web ���ε{���C
    </ul>
  <p>
  <span class="clsEntryText">
    IIS �N���{�� Internet �зǩM Windows ��X�b�@�_�A�o�ˡA�ϥ� Web �ä�����
    �n���Y�}�l�ǲ߷s���覡�ӵo��B�޲z�ζ}�o���e�C
  <p>
  </span>
  </td>

    <td valign="top">
      <img src="mmc.gif">
    </td>
    <td valign="top">
      <span class="clsHeading">��X���޲z</span>
      <br>
      <span class="clsEntryText">
        �z�i�H�z�L�U�C�u��Ӻ޲z IIS: Windows XP [�q���޲z] <a href="javascript:activate(<%=iVer%>);">�D���x</a> 
        �ΨϥΫ��O�X�C�ϥΥD���x�A�z�]�i�H�z�L Web �N���x�Φ��A�� (�� IIS �Һ޲z) �����e�@�ε���L�H�C
        �q�D���x�s�� IIS �O�J���޲z�椸�A�z�i�H
        �]�w�̱`�Ϊ� IIS �]�w�Τ��e�C�b���x�����ε{���}�o����A�o�ǳ]�w�Τ��e�i�H�Φb
        ������¤O�� Windows ���A���������Ͳ����Ҥ��C 
      <p>
       
      </span>
    </td>
  </tr>
  <tr>
    <td valign="top">
      <img src="help.gif">
    </td>
    <td valign="top">
      <span class="clsHeading"><a href="javascript:loadHelpFront();">�u�W���</a></span>
      <br>
      <span class="clsEntryText">IIS �u�W���]�t���ޡB�����˯�
        �Ψ̸`�I�έӧO�D�D�C�L����O�C���{���]�p�޲z�Ϋ��O�X
        �޲z�A�Шϥ� IIS �Ҵ��Ѫ��d�ҡC�����ɮ׷|�s��
        HTML�A���\�z���ݭn���@���ѤΦ@�ΡC�ϥ� IIS �u�W���A
        �z�i�H:<p>
      </span>
      <ul class="clsEntryText">
         <li>���o�u�@����
         <li>�ǲߦ��A���ާ@�κ޲z
         <li>�d�\�ѦҸ��
         <li>�˵��{���X�d�ҡC
      </ul>
      <p>
        <span class="clsEntryText">
        �䥦���� IIS �����Τά�������T�ӷ���� Microsoft.com 
        ����: MSDN�BTechNet �� Windows ���x�C
        </span>
    </td>
  </tr>
  
  <tr>
    <td valign="top">
      <img src="print.gif">
    </td>
    <td valign="top">
      <span class="clsHeading">Web �C�L</span>
      <br>
      <span class="clsEntryText">Windows XP Professional �|�ʺA�C�X�Ҧ����A�� (�b�i����
        �s���������W) �W���L�����C�z�i�H�s�������x��
        �ʵ��L�����Ψ�u�@�C�z�]�i�H�q���� Windows �q���z�L�����x�s����L�����C
        �аѾ\���� Internet �C�L�� Windows �������C
      </span>
    </td>
  </tr>
  
  </table>
</td>
</tr>
</table>

<p align=center><em><a href="/iishelp/common/colegal.htm">c 1997-2001 Microsoft Corporation. All rights reserved.</a></em></p>

</body>
</html>

<% End If %>
