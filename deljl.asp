<%
if session("userright")<>"�ֿ�����" and session("userright")<>"�ֹ�" then
   response.write "��û��ʹ�õ�Ȩ�ޣ�"
   response.end
end if
%>
<!-- #include file="data\config.asp" -->
<%
conn.execute "delete from UnitInfo where UnitName='" & request("jln") & "'"
conn.close
set conn=nothing
success="������λɾ���ɹ����ȴ�ϵͳ�Զ�����......"
url="jiliang_list.asp"
%>

<html><head><title>ɾ��������λ</title><meta http-equiv=refresh content="2; url=<%=url%>">
<table width=100% height=100%>
<tr>
<td align=center>
<b><%=success%></b>
</td>
</tr>
</table>
</html>