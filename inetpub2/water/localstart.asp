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

<title>歡迎使用 Windows XP Server Internet 服務</title>
<style>
  ul{margin-left: 15px;}
  .clsHeading {font-family: 新細明體; color: black; font-size: 11; font-weight: 800; width:210;}  
  .clsEntryText {font-family: 新細明體; color: black; font-size: 11; font-weight: 400; background-color:#FFFFFF;}    
  .clsWarningText {font-family: 新細明體; color: #B80A2D; font-size: 11; font-weight: 600; width:550;  background-color:#EFE7EA;}  
  .clsCopy {font-family: 新細明體; color: black; font-size: 11; font-weight: 400;  background-color:#FFFFFF;}  
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
  <b>Web 服務現在執行中。
  
<% If Not bSuccess And Err = 0 Then %>
  
  <p>您目前並沒有為使用者建立一個預設 Web 頁面。
任何嘗試從另一台電腦連接您網站的使用者目前會收到一個顯示
  <a href="iisstart.asp?uc=1">建構中</a> 的頁面。
  您的 Web 伺服器列出下列檔案為可能的預設 Web 頁面: <%=sDefDoc%>。目前只有 iisstart.asp 存在。<br><br>
  
<% End If %>

  若要新增文件到預設網站，請將檔案儲存在 <%=sPhyspath%>。
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
  歡迎使用 IIS 5.1</span><br>
      <span class="clsEntryText">    
    Internet Information Services (IIS) 5.1 for Microsoft Windows XP Professional
    將網路運算的威力 
    帶到了 Windows。有了 IIS，您可以輕易地共用檔案及印表機，或者您可以建立應用程式，
    在 Web 上安全地發行資訊，以增進您組織共用資訊的方式。IIS 是一個安全的平台，
    適合用來建立及調配電子商務解決方案，以及重要的 Web 應用程式。
  <p>
    使用已安裝 IIS 的 Windows XP Professional，提供一個個人及開發的作業系統，讓您可以:</span>
  <p>
    <ul class="clsEntryText">
      <li>安裝個人的 Web 伺服器
      <li>在小組中共用資訊
      <li>存取資料庫
      <li>開發企業內部網路
      <li>開發 Web 應用程式。
    </ul>
  <p>
  <span class="clsEntryText">
    IIS 將公認的 Internet 標準和 Windows 整合在一起，這樣，使用 Web 並不表示
    要重頭開始學習新的方式來發行、管理或開發內容。
  <p>
  </span>
  </td>

    <td valign="top">
      <img src="mmc.gif">
    </td>
    <td valign="top">
      <span class="clsHeading">整合的管理</span>
      <br>
      <span class="clsEntryText">
        您可以透過下列工具來管理 IIS: Windows XP [電腦管理] <a href="javascript:activate(<%=iVer%>);">主控台</a> 
        或使用指令碼。使用主控台，您也可以透過 Web 將站台及伺服器 (由 IIS 所管理) 的內容共用給其他人。
        從主控台存取 IIS 嵌入式管理單元，您可以
        設定最常用的 IIS 設定及內容。在站台及應用程式開發之後，這些設定及內容可以用在
        執行更具威力的 Windows 伺服器版本的生產環境中。 
      <p>
       
      </span>
    </td>
  </tr>
  <tr>
    <td valign="top">
      <img src="help.gif">
    </td>
    <td valign="top">
      <span class="clsHeading"><a href="javascript:loadHelpFront();">線上文件</a></span>
      <br>
      <span class="clsEntryText">IIS 線上文件包含索引、全文檢索
        及依節點或個別主題列印的能力。對於程式設計管理及指令碼
        管理，請使用 IIS 所提供的範例。說明檔案會存放為
        HTML，允許您視需要對其作註解及共用。使用 IIS 線上文件，
        您可以:<p>
      </span>
      <ul class="clsEntryText">
         <li>取得工作說明
         <li>學習伺服器操作及管理
         <li>查閱參考資料
         <li>檢視程式碼範例。
      </ul>
      <p>
        <span class="clsEntryText">
        其它有關 IIS 的有用及相關的資訊來源位於 Microsoft.com 
        網站: MSDN、TechNet 及 Windows 站台。
        </span>
    </td>
  </tr>
  
  <tr>
    <td valign="top">
      <img src="print.gif">
    </td>
    <td valign="top">
      <span class="clsHeading">Web 列印</span>
      <br>
      <span class="clsEntryText">Windows XP Professional 會動態列出所有伺服器 (在可輕易
        存取的網站上) 上的印表機。您可以瀏覽此站台來
        監視印表機及其工作。您也可以從任何 Windows 電腦透過此站台連接到印表機。
        請參閱有關 Internet 列印的 Windows 說明文件。
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

