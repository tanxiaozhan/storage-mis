<!-- #include file="data\config.asp" -->
<html><title>����ǰ׺�޸�</title>
<%
    Set rs =server.createobject("adodb.recordset")
	
	ltimeout=Server.ScriptTimeOut 
	Server.ScriptTimeOut =50000
	%>
	<script language="vbscript">
		dim choice	
		Choice = msgbox ("��ʼ�޸�����ǰ׺��?",vbYESNO,"��ʾ")
		if Choice = vbNO then
			history.back()
		end if
	</script>
	
	<%

	'�޸����ֵ���ǰ׺
	sql = "Update InStore " & _
	      "set BillNo ='J' + BillNo" & " " & _
		  "where  left(BillNo,1) <> 'J' "
	'response.write(sql)
	
		  
	conn.Execute(sql)
	
	'�޸����ֵ���ǰ׺
	sql = "Update OutStore " & _
	      "set BillNo ='C' + BillNo" & " " & _
		  "where  left(BillNo,1) <> 'C' "
	'response.write(sql)
	
		  
	conn.Execute(sql)
	
	
	
	conn.close
	set conn=nothing

	Server.ScriptTimeOut =ltimeout
	
	info="�����޸��ɹ�!2���ϵͳ�Զ����ء�"
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