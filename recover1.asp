<%
if session("userright")<>"仓库主管" then
   response.write "你没有使用的权限！"
   response.end
end if
%>

<%

  if request("tj")="恢复" and  trim(request("filelist"))<>"" then
     set fs=server.createobject("scripting.filesystemobject")
   
     'response.write server.mappath("back\" & trim(request("filelist")))
     if fs.fileexists(server.mappath("back\" & trim(request("filelist")))) then
        fs.copyfile server.mappath("back\" & trim(request("filelist"))),server.mappath("data\SMIS.mdb")
     else
        response.write "文件不存在！[<a href='javascript:history.back()'>返回</a>]"
        response.end 
     end if
            
  else
  
     response.write "文件不存在！[<a href='javascript:history.back()'>返回</a>]"
     response.end
     
  end if

success="文件恢复成功！等待系统自动返回......"
url="recover.asp"

%>
<html><head><title>客户信息</title><meta http-equiv=refresh content="1; url=<%=url%>">
<table width=100% height=100%>
<tr>
<td align=center>
<b><%=success%></b>
</td>
</tr>
</table>
</html>