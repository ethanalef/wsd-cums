<!-- #include file="../conn.asp" -->

<%
SQl = "SELECT  a.memno,a.adate,sum(a.bankin),b.memname,b.memcname  FROM  autopay a ,memmaster b where a.memno=b.memno and a.flag='F' and right(a.code,1)='1' group by a.memno,a.adate,b.memname,b.memcname order by a.memno,a.adate,b.memname,b.memcname "
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn,1,1
if not rs.eof then 
ttlcnt=0 
payamt = 0


	spaces=""
	for idx = 1 to 50
		spaces=spaces&" "
	next
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.CreateTextFile(Server.MapPath("..\txt")&"\autorejlst.txt", True)
        memno = 0
      
        
        mx = 0
        do while mx = 0  and not rs.eof
 

           if memno <> rs("memno")  then
              payamt = rs(2)
              for i= 1 to 6
              objFile.WriteLine ""
              next
              
              objFile.Write "檔案編號 Ref# : "
              objFile.WriteLine ""
              objFile.WriteLine ""
              objFile.Write "　　　　　親愛的社員  姓名 "
              objFile.Write rs("memcname") 
              objFile.WriteLine ""
              objFile.WriteLine ""
              objFile.Write "　　　　　社員編號："    
              objFile.Write  rs("memno")
              objFile.WriteLine ""
              objFile.WriteLine ""	
              objFile.Write  left(spaces,20)
              objFile.Write  "　　　　　　　　　　　　　　　　　　　按月銀行自動轉帳失效通知書"
              objFile.WriteLine ""
              objFile.WriteLine ""
              objFile.Write  "　　　　　　　　　　多謝你一直以來對本社的支持和信任！"
              objFile.WriteLine ""
              objFile.WriteLine ""
              
              objFile.Write "　　　　　　　　　　依據本社記錄顯示，從銀行覆函得知，本社於上月底未能從閣下戶口收取" 
              objFile.WriteLine ""
              objFile.Write "　　　　　該月應繳納之款項，相信可能是一時忘記。請閣下於接獲此通知書後，從速聯絡本"
              objFile.WriteLine ""
              objFile.Write "　　　　　社辦事處或各分區委員安排補回款項 $"  
              objFile.Write formatnumber(payamt,2)
              objFile.Write " 。敝社定當盡力協助閣下解決有關"
              objFile.WriteLine ""
              objFile.Write  "　　　　　財務事宜，以免影響日後借貸信用或其他利息上之損失。"     
              objFile.WriteLine ""
              objFile.WriteLine ""
          
              objFile.Write  "　　　　　本社不鼓勵社員經常發生自動轉帳和還款脫期等情況；而為了加強本社之"
              objFile.WriteLine "" 
              objFile.Write  "　　　　　工作效率與及保障其他社員權益，我們會盡量在不滋擾閣下工作的情況下，經常及"
              objFile.WriteLine ""
              objFile.Write  "　　　　　第一時間提醒閣下與及閣下之擔保人（如適用）有關上述事宜，不便之處，敬請原"
              objFile.WriteLine ""
              objFile.Write  "　　　　　諒！" 
              objFile.WriteLine ""
              objFile.WriteLine ""
              objFile.Write  left(spaces,4)
              objFile.Write "　　　　　假如閣下已從其他方式，例如現金、支票或過戶形式，補回上述之款項，"
              objFile.WriteLine "" 
              objFile.Write  "　　　　　則無須理會此通知書。" 
              objFile.WriteLine "" 
              objFile.Write  left(spaces,4)
              objFile.Write "　　　　　如有任何查詢，歡迎致電2787 9222與本社職員聯絡。，"
              objFile.WriteLine "" 
              objFile.WriteLine "" 
              objFile.WriteLine "" 
              objFile.WriteLine "" 
              objFile.Write  "　　　　　水務署員工儲蓄互助社"
              objFile.WriteLine "" 
              objFile.Write  "　　　　　董事會 司庫"
              objFile.WriteLine "" 
　　　　　　　objFile.Write  "　　　　　”　　　　　
              objFile.Write rs("adate")  
              objFile.WriteLine "" 
              objFile.WriteLine "" 
              objFile.Write  "　　　　　副本呈交"
              objFile.WriteLine "" 
              objFile.Write  "        呆帳小組"
              objFile.WriteLine "" 
              objFile.Write  "        貸款委員會"
              objFile.WriteLine ""  
              objFile.WriteLine ""  
              objFile.Write "　　　　　本函乃由電腦集中印發，無須簽署"
              objFile.WriteLine ""  
              for i = 1 to 26
              objFile.WriteLine ""
              next
             
              payamt = 0
            
              memno=rs("memno")
             
        end if
        if not rs.eof then 
           rs.movenext
        else
           mx = 1
        end if
        loop
      
	objFile.Close

	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.redirect "../txt/autorejlst.txt"
end if
%>
<html>
<head>
<title>銀行自動轉帳失效通知書</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
</body>
<html>



