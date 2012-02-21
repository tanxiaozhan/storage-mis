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
	
	

		'关闭数据库连接
		conn.close
		set conn=nothing
		
   		response.redirect  "deleteoutstore.asp"
   		
   		
   end if
%>

<html><head><title>进仓管理</title>
<p align="center"><b><font size="5">恢复出仓记录</font></b></p>
<p>　</p>
<p></p>
<form method="POST" action="recoverout.asp?RecNo=<% = request("recno") %>">
  <p>请输入与之相关的进仓单记录号:<input type="text" name="tinrecno" size="9"><input type="submit" value="恢复" name="B1"><input type="button" value="返回" name="B2" onclick="vbscript:history.back()"></p>
</form>
</html>
