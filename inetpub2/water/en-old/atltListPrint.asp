<!-- #include file="../conn.asp" -->

<%
SQl = "SELECT  a.memno,a.adate,sum(a.bankin),b.memname,b.memcname,b.accode  FROM  autopay a ,memmaster b where a.memno=b.memno and a.pflag=1 and a.status='F' and right(a.code,1)='1' group by a.memno,a.adate,b.memname,b.memcname,b.accode order by a.memno,a.adate,b.memname,b.memcname,b.accode "
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn,1,1

if not rs.eof then 
ttlcnt=0 
payamt = 0
zz = 1

	spaces=""
	for idx = 1 to 50
		spaces=spaces&" "
	next
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.CreateTextFile(Server.MapPath("..\txt")&"\autorejlst.txt", True)
        memno = 0
      
        
        mx = 0
        do while not rs.eof
 

         
              refno = "AR"&right(year(date()),2)&right("0"&month(date())-1,2)&right("0"&zz,2)
              payamt = rs(2)
              for i = 1 to 6
                  objFile.WriteLine ""
              next
              objFile.Write "�ɮ׽s���@Ref# : "
               objFile.Write refno
              objFile.WriteLine ""
              objFile.WriteLine ""
              objFile.Write "�@�@�@�@�@�˷R������  �m�W "
              objFile.Write rs("memcname") 
              objFile.WriteLine ""
              objFile.WriteLine ""
              objFile.Write "�@�@�@�@�@�����s���G"    
              objFile.Write  rs("memno")
              objFile.WriteLine ""
              objFile.WriteLine ""	
              
              objFile.Write  "�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@����Ȧ�۰���b���ĳq����"
              objFile.WriteLine ""
              objFile.Write  "�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@--------------------------"
              objFile.WriteLine ""
              objFile.Write  "�@�@�@�@�@�@�@�@�@�h�§A�@���H�ӹ糧��������M�H���I"
              objFile.WriteLine ""
              objFile.WriteLine ""
              
              objFile.Write "�@�@�@�@�@�@�@�@�@�̾ڥ����O����ܡA�q�Ȧ��Ш�o���A������W�멳����q�դU��f����" 
              objFile.WriteLine ""
              objFile.Write "�@�@�@�@�@�Ӥ���ú�Ǥ��ڶ��A�۫H�i��O�@�ɧѰO�C�лդU���򦹳q���ѫ�A�q�t�p����"
              objFile.WriteLine ""
              objFile.Write "�@�@�@�@�@����ƳB"
              if rs("accode")<>9999 then
              objFile.Write "�ΦU���ϩe���w�Ƹɦ^�ڶ� $"                          
              objFile.Write formatnumber(payamt,2)
              objFile.Write " �C�ͪ��w��ɤO��U�դU�ѨM����"
              objFile.WriteLine ""
              objFile.Write  "�@�@�@�@�@�]�ȨƩy�A�H�K�v�T���ɶU�H�ΩΨ�L�Q���W���l���C"     
              objFile.WriteLine ""
              else
              objFile.Write "�w�Ƹɦ^�ڶ� $"                          
              objFile.Write formatnumber(payamt,2)
              objFile.Write " �C�ͪ��w��ɤO��U�դU�ѨM�����]�ȨƩy�A�H�K"
              objFile.WriteLine ""
              objFile.Write  "�@�@�@�@�@�v�T���ɶU�H�ΩΨ�L�Q���W���l���C"     
              objFile.WriteLine ""              
              end if
              objFile.WriteLine ""
             
              objFile.Write "�@�@�@�@�@�@�@�@�@���������y�����g�`�o�ͦ۰���b�M�ٴڲ�������p�F�Ӭ��F�[�j������"
              objFile.WriteLine "" 
              objFile.Write  "�@�@�@�@�@�u�@�Ĳv�P�ΫO�٨�L�����v�q�A�ڭ̷|�ɶq�b�����Z�դU�u�@�����p�U�A�g�`��"
              objFile.WriteLine ""
              objFile.Write "�@�@�@�@�@�Ĥ@�ɶ������դU�P�λդU����O�H�]�p�A�Ρ^�����W�z�Ʃy�A���K���B�A�q�Э�"
              objFile.WriteLine ""
              objFile.Write  "�@�@�@�@�@�̡I" 
              objFile.WriteLine ""
              objFile.WriteLine ""
             
              objFile.Write "�@�@�@       �@�@���p�դU�w�q��L�覡�A�Ҧp�{���B�䲼�ιL��Φ��A�ɦ^�W�z���ڶ��A"
              objFile.WriteLine "" 
              objFile.Write  "�@�@�@�@�@�h�L���z�|���q���ѡC" 
              objFile.WriteLine "" 
              objFile.WriteLine ""
              objFile.Write "�@       �@�@�@�@�p������d�ߡA�w��P�q2787 9222�P����¾���p���C�A"
              objFile.WriteLine "" 
              objFile.WriteLine "" 
              objFile.WriteLine "" 
              objFile.WriteLine "" 
              objFile.Write  "�@�@�@�@�@���ȸp���u�x�W���U��"
              objFile.WriteLine "" 
              objFile.Write  "�@�@�@�@�@���Ʒ| �q�w"
              objFile.WriteLine "" 
              objFile.Write  "�@�@�@�@�@"
              objFile.Write year(date())
              objFile.Write  "�~"
              objFile.Write month(date())
              objFile.Write  "��"
              objFile.Write day(date())
              objFile.Write  "��"
              objFile.WriteLine "" 
              objFile.WriteLine "" 
              objFile.Write  "�@�@�@�@�@�ƥ��e��"
              objFile.WriteLine "" 
              objFile.Write  " �@�@�@�@�@       �b�b�p��"
              objFile.WriteLine "" 
              objFile.Write  "     �@�@�@�@�@   �U�کe���|"
              objFile.WriteLine ""  
              objFile.WriteLine ""  
              objFile.Write "�@�@�@�@�@����D�ѹq�������L�o�A�L��ñ�p"
              objFile.WriteLine ""  
              for i = 1 to 25
              objFile.WriteLine ""
              next
              payamt = 0
            
             
             
      
 
        zz=zz+1
         rs.movenext  
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
<title>�Ȧ�۰���b���ĳq����</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
</body>
<html>



