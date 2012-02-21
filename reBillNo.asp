<!-- #include file="data\config.asp" -->
<html><title>单号前缀修复</title>
<%
    Set rs =server.createobject("adodb.recordset")
	
	ltimeout=Server.ScriptTimeOut 
	Server.ScriptTimeOut =50000
	%>
	<script language="vbscript">
		dim choice	
		Choice = msgbox ("开始修复单号前缀吗?",vbYESNO,"提示")
		if Choice = vbNO then
			history.back()
		end if
	</script>
	
	<%

	'修复进仓单号前缀
	sql = "Update InStore " & _
	      "set BillNo ='J' + BillNo" & " " & _
		  "where  left(BillNo,1) <> 'J' "
	'response.write(sql)
	
		  
	conn.Execute(sql)
	
	'修复出仓单号前缀
	sql = "Update OutStore " & _
	      "set BillNo ='C' + BillNo" & " " & _
		  "where  left(BillNo,1) <> 'C' "
	'response.write(sql)
	
		  
	conn.Execute(sql)
	
	
	
	conn.close
	set conn=nothing

	Server.ScriptTimeOut =ltimeout
	
	info="单号修复成功!2秒后系统自动返回。"
    url="jiliang_list.htm"
%>

<meta http-equiv=refresh content="2; url=<%=url%>">


<table width=100% height=100%>
<tr>
<td align=center>
<font size="6">
<b><%=info%></b></font>
</td>
</tr>
</table>
</html>