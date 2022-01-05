<!-- #include file="../conn.asp" -->
<% ' <!-- #include file="../CheckUserStatus.asp" --> %>
<%


server.scripttimeout = 1800

         set rs = conn.execute("select sum(bankin) from autopay where right(code,1)='1'  " )
         if not rs.eof then
            sumttl = rs(0)*100
         end if
         rs.close
         cnt = 0
         set rs = conn.execute("select  memno  from autopay where right(code,1)='1'  group by memno order by memno " )
         do while  not rs.eof 
            cnt = cnt + 1
            rs.movenext
         loop
         rs.close
 
	 set rs = server.createobject("ADODB.Recordset")
         sql  = "select a.memno,sum(a.bankin),a.adate,b.memcname,b.memname,b.bnk,b.bch,b.bacct from autopay a ,memmaster b where a.memno=b.memno and right(a.code,1)='1'  group by a.memno,adate,b.memcname,b.memname,b.bnk,b.bch,b.bacct order by a.memno,adate,b.memcname,b.memname,b.bnk,b.bch,b.bacct "
         rs.open sql, conn,1,1
         if rs.eof then
 
            response.redirect "menu.asp" 
         end if



ttlamt = 0
        xday = right("0"&day(rs("adate")),2)
        xmon = right("0"&month(rs("adate")),2)
        xyr  = right(year(rs("adate")),2)
        yr = right(xyr,2)
        mn  = month(date())
        xmn=  mid("JANFEBMARAPRMAYJUNJULAUGSEPOCTNOVDEC",(mn-1)*3+1,3)

        xcnt= right("00000"&cnt,5)        
        xdate = xday&xmon&xyr
        
 
	spaces=""
	for idx = 1 to 50
		spaces=spaces&" "
	next

        xmark1 = left(xmn&" "&yr&spaces,12)
        xttl = right("000000000"&sumttl,10)


	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.CreateTextFile("c:\public\autopay.apc", True)

	objFile.Write "G"
        objFile.Write "024062010001"
        objFile.Write "R01"
        objFile.Write xmark1
        objFile.Write xdate   
	objFile.Write "K********"      
        objFile.Write xcnt 
        objFile.Write  xttl

        objFile.Write  "                     1"
        xx = 0
 do while  NOT rs.eof                 
        xx = xx + 1
      if rs("memno")="4660" or rs("memno")="3568" or rs("memno")="2453" or rs("memno")="2580" or rs("memno")="4457" or rs("memno")="4318" then
            
          memno = right("     "&rs("memno"),4) 
        else   
          memno = right("     "&rs("memno"),5) 
       end if  
        memname = rs("memname")
        pos = instr(memname,", ")
        if pos > 0 then
           memname = left(rs("memname"),pos)&mid(rs("memname"),pos+2,len(rs("memname"))-pos)            
        end if
       if len(memname)>= 20 then
          memname = left(memname,20)
       end if
       xbnk = rs("bnk")&rs("bch")&rs("bacct")
       ln =  len(xbnk)
       xbnk = xbnk&left("              ",15)

        objFile.Write " NO"
 
            objFile.Write  left(memno&spaces,10)
      
        objFile.Write  left(memname&spaces,20)  
         objFile.Write left(xbnk&spaces,15)
 
       
        samt = round(rs(1),2)*100
        nsamt = right("0000000000"&samt,10)   
        objFile.Write  nsamt
        objFile.Write  left(spaces,22)
        
      
    rs.movenext
    loop
   objFile.Close

	set rs=nothing
	conn.close
	set conn=nothing


  
    RESPONSE.REDIRECT "COMPLETED.ASP"

 

%>
<html>
<head>
<title></title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">



</body>
</html>

