<!-- #include file="data\config.asp" -->
<%

  Set rs =server.createobject("adodb.recordset")
  sql="select * from master"
  rs.Open sql,conn,3,3
  
  if rs.recordcount>0 then
		rs("Name") = request("t2")
		rs("Address") = request("t3")
		rs("Tel") = request("t4")
		rs("Fax") = request("t5")
  	    rs.update
  	    success="��˾��Ϣ�޸ĳɹ����ȴ�ϵͳ�Զ�����......"

  else
  		response.write("����ʧ�ܣ������²�����")
  end if
  
  rs.close
  set rs=nothing
  conn.close
  set conn=nothing
url="jiliang_list.htm" 

%>

<html><head><title>��Ӽ�����λ</title><meta http-equiv=refresh content="0.5; url=<%=url%>">


<table width=100% height=100%>
<tr>
<td align=center>
<b><%=success%></b>
</td>
</tr>
</table>
</html>