<%
function jiami(strf)
   
   randomize
   tmp=""
   
   sz=(rnd * 25) + 90'65
   for i=1 to len(strf)
       tmp=tmp & chr((asc(mid(strf,i,1)) xor sz))
   next
   
   jiami=tmp & chr(sz)
end function


function jiemi(strf)
   
   sz=asc(right(strf,1))
   
   tmp=""
   
   for i=1 to len(strf)-1
       tmp=tmp & chr((asc(mid(strf,i,1)) xor sz))
   next
   
   jiemi=tmp
      
end function

%>

<%
    dim conn
    dim  connectstr
    set  conn=server.createobject("adodb.connection")
    connectstr="provider=microsoft.jet.oledb.4.0; jet oledb:database password=office;data source=" & server.mappath("data\SMIS.mdb")
    conn.Open connectstr 

 sql="select * from userInfo where user='" & request("us") & "'"
 
 Set rs =server.createobject("adodb.recordset")
 rs.Open sql,conn,3,3
 
 if request("mm")<>trim(jiemi(rs("pwd"))) then
    rs.close
    set rs=nothing
    conn.close
    set conn=nothing
    response.write "密码输入错误！[<a href='javascript:history.back()'>返回</a>]"
    response.end
 else
    session("username")=rs("user") 
    session("userright")=rs("right")
    
    set fs=server.createobject("scripting.filesystemobject")
    'set fs1=server.createobject("scripting.filesystemobject")


    if fs.folderexists(server.mappath("back"))=false then fs.createfolder(server.mappath("back"))
    
    if fs.GetFolder(server.mappath("back")).Files.Count>=20 then 
      
      a=""
      for each fs1 in fs.getfolder(server.mappath("back")).files
          if a="" then 
             a=fs1.path
          else
             if fs1.path < a then '得到最早的备份文件名
                a=fs1.path
             end if
          end if
      next
      if a<>"" then fs.deletefile(a)
    end if   
         
    if fs.fileexists(server.mappath("data\SMIS.mdb")) then
         path1="back\" & year(now()) & "_" & month(now()) & "_" & day(now()) & "_" & hour(now()) & "_" & minute(now()) & "_" & second(now()) & ".bak"
         fs.copyfile server.mappath("data\SMIS.mdb"),server.mappath(path1)
         'response.write server.mappath(path1)
         'response.end
    end if
     
    response.redirect "show.asp"
    
 end if
%>
<%
'<SCRIPT language=javascript> window.open('hpzz/sj.htm','考试时间','scrollbars=yes,width=600,height=400,fullscreen=yes')</SCRIPT>
 'setTimeout("self.close()",50000) 

'<SCRIPT language=javascript>
'   window.open('Client_list.htm','仓库管理系统','scrollbars=yes,resizable=yes')
   
'</SCRIPT>
%>