<!-- #include file="data\config.asp" -->

<%
	if trim( request("tinrecno") )<>"" then
	
		sql="Update delOutStore " & _
		    "set InstoreRecNo=" & request("tinrecno") & " " & _
	    	"where RecNo=" & request("RecNo")
	
		conn.Execute(sql)
	
	
		sql="insert into outstore " & _
	    	"select * from deloutstore " & _
		    "where RecNo=" & request("RecNo")
	
		conn.execute(sql)
	
	
		sql="delete from deloutstore " & _
		    "where recno=" & request("recno")
		conn.execute(sql)
	
	

		'�ر����ݿ�����
		conn.close
		set conn=nothing
		
   		response.redirect  "deleteoutstore.asp"
   		
   		
   end if
%>

<html><head><title>���ֹ���</title>
<p align="center"><b><font size="5">�ָ����ּ�¼</font></b></p>
<p>��</p>
<p></p>
<form method="POST" action="recoverout.asp?RecNo=<% = request("recno") %>">
  <p>��������֮��صĽ��ֵ���¼��:<input type="text" name="tinrecno" size="9"><input type="submit" value="�ָ�" name="B1"><input type="button" value="����" name="B2" onclick="vbscript:history.back()"></p>
</form>
</html>
