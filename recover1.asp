<%
if session("userright")<>"�ֿ�����" then
   response.write "��û��ʹ�õ�Ȩ�ޣ�"
   response.end
end if
%>

<%

  if request("tj")="�ָ�" and  trim(request("filelist"))<>"" then
     set fs=server.createobject("scripting.filesystemobject")
   
     'response.write server.mappath("back\" & trim(request("filelist")))
     if fs.fileexists(server.mappath("back\" & trim(request("filelist")))) then
        fs.copyfile server.mappath("back\" & trim(request("filelist"))),server.mappath("data\SMIS.mdb")
     else
        response.write "�ļ������ڣ�[<a href='javascript:history.back()'>����</a>]"
        response.end 
     end if
            
  else
  
     response.write "�ļ������ڣ�[<a href='javascript:history.back()'>����</a>]"
     response.end
     
  end if

success="�ļ��ָ��ɹ����ȴ�ϵͳ�Զ�����......"
url="recover.asp"

%>
<html><head><title>�ͻ���Ϣ</title><meta http-equiv=refresh content="1; url=<%=url%>">
<table width=100% height=100%>
<tr>
<td align=center>
<b><%=success%></b>
</td>
</tr>
</table>
</html>