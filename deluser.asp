<%
if session("userright")<>"�ֿ�����" and session("userright")<>"�ֹ�" then
   response.write "��û��ʹ�õ�Ȩ�ޣ�"
   response.end
end if
%>
<!-- #include file="data\config.asp" -->
<%
conn.execute "delete from ClientInfo where id=" & request("id")
conn.close
set conn=nothing
success="�ͻ���Ϣɾ���ɹ����ȴ�ϵͳ�Զ�����......"
url="client_list.asp"
%>

<html><head><title>ɾ���ͻ���Ϣ</title><meta http-equiv=refresh content="2; url=<%=url%>">
<table width=100% height=100%>
<tr>
<td align=center>
<b><%=success%></b>
</td>
</tr>
</table>
</html>