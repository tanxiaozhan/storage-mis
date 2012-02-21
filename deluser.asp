<%
if session("userright")<>"仓库主管" and session("userright")<>"仓管" then
   response.write "你没有使用的权限！"
   response.end
end if
%>
<!-- #include file="data\config.asp" -->
<%
conn.execute "delete from ClientInfo where id=" & request("id")
conn.close
set conn=nothing
success="客户信息删除成功！等待系统自动返回......"
url="client_list.asp"
%>

<html><head><title>删除客户信息</title><meta http-equiv=refresh content="2; url=<%=url%>">
<table width=100% height=100%>
<tr>
<td align=center>
<b><%=success%></b>
</td>
</tr>
</table>
</html>