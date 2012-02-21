<%
if session("userright")<>"仓库主管" then
   response.write "你没有使用的权限！【<a href='javascript:history.back()'>返回</a>】"
   response.end
end if
%>
<!-- #include file="data\config.asp" -->
<%
conn.execute "delete from userInfo where user='" & request("usern") & "'"
conn.close
set conn=nothing
success="用户删除成功！等待系统自动返回......"
url="user_list.asp"
%>

<html><head><title>删除用户</title><meta http-equiv=refresh content="2; url=<%=url%>">
<table width=100% height=100%>
<tr>
<td align=center>
<b><%=success%></b>
</td>
</tr>
</table>
</html>