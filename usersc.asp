<%
if session("userright")<>"�ֿ�����" then
   response.write "��û��ʹ�õ�Ȩ�ޣ���<a href='javascript:history.back()'>����</a>��"
   response.end
end if
%>
<!-- #include file="data\config.asp" -->
<%
conn.execute "delete from userInfo where user='" & request("usern") & "'"
conn.close
set conn=nothing
success="�û�ɾ���ɹ����ȴ�ϵͳ�Զ�����......"
url="user_list.asp"
%>

<html><head><title>ɾ���û�</title><meta http-equiv=refresh content="2; url=<%=url%>">
<table width=100% height=100%>
<tr>
<td align=center>
<b><%=success%></b>
</td>
</tr>
</table>
</html>