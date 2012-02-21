<!-- #include file="data\config.asp" -->
<html><title>数据修复</title>
<%
	
	response.write("<BR><BR><BR><center>此功能暂停!<BR><BR>")
	response.write("<a href='jiliang_list.htm'>返回</a></center>")
	response.end
    Set rs =server.createobject("adodb.recordset")
	
	ltimeout=Server.ScriptTimeOut 
	Server.ScriptTimeOut =50000
	'进仓结余数据恢复为初值
	%>
	
	
	<script language="vbscript">
		dim choice	
		Choice = msgbox ("开始进行数据修复吗?",vbYESNO,"提示")
		if Choice = vbNO then
			history.back()
		else
			window.alert "准备初始化进仓结余数据!"
		end if
	</script>
	
	<%

	sql = "Update InStore " & _
	      "set remainnumber = allnumber," & _
	      "remainpack = packnumber," & _
	      "remainweight=weight"
	
	conn.Execute(sql)
	
	response.write("初始化进仓结余数据成功!<BR><BR>")

	

	'删除出仓数据库中的残余记录

	%>
	<script language="vbscript">
		window.alert "准备修复出仓冗余数据!"
	</script>
	
	<%


	dim lRecno(1000) 
	dim n 
	n=0
	set instoreRS =server.createobject("adodb.recordset")
	set outstoreRS=server.createobject("adodb.recordset")
	sql ="select recno,InstoreRecNo from outstore"
	outstoreRS.open sql,conn,1,1
	do while not outstoreRS.eof

		sql="select recno from instore where recno = " & cstr( outstoreRS("instoreRecNo") )
		instoreRS.open sql,conn,1,1
		if instoreRs.RecordCount < 1 then
			n=n+1
			lRecno(n)=outstoreRS("instoreRecNo")	
		end if
		instoreRS.close
		outstorers.movenext
	
	loop
	
	outstoreRS.close
	
	for i= 1 to n
		sql="delete from outstore where instoreRecNo=" & cstr( lrecno(i) )
		conn.execute(sql)
	next 	
	response.write("修复出仓冗余数据成功!<BR><BR>")


	'用进仓数据库中的数据更新出仓库的数据

	%>
	<script language="vbscript">
		window.alert "准备更新出仓数据!"
	</script>
	
	<%
	
	sql="select * from instore"
	instoreRS.open sql,conn,1,1

	do while not instoreRS.eof
		sql = "Update OutStore " & _
		      "set ClientID = '" & instoreRS("ClientID") & "'," & _
	    	  "contractID='" & instoreRS("contractID") & "'," & _
	      	  "Gross=" & cstr(instoreRS("Gross")) & "," & _
	      	  "net=" & cstr(instoreRS("net")) & "," & _
	      	  "PackWeight=" & cstr(instoreRS("PackWeight") ) & " " & _
	      	  "where instoreRecno=" & instoreRS("RecNo")
	      	  

		conn.execute(sql)
		
		'重量未锁定的,重量=包重*总件数
		sql = "Update OutStore " & _     
	      	  "set weight=pack1*" & cstr( instoreRS("packweight") ) & "/1000  " & _ 
	      	  "where instoreRecno=" & instoreRS("RecNo") & " and lockweight=FALSE"

		conn.execute(sql)				

		instoreRS.movenext
	
	loop

	response.write("更新出仓数据成功!<BR><BR>")



	'根据出仓数量,修复进仓结余数据
	sql = "select * from OutStore"
	rs.open sql,conn,1,1
	
	%>
	<script language="vbscript">
		window.alert "准备修复进仓结余数据!"
	</script>
	
	<%

	do while not rs.eof
		
		sql = "Update InStore " & _
		      "set remainnumber = remainnumber - " & rs("number1") & "," & _
	    	  "remainpack = remainpack - " & rs("pack1") & "," & _
	      	  "remainweight=remainweight - " & rs("weight") & " " & _
	      	  "where Recno=" & rs("InStoreRecNo")
	    'response.write(sql)
	    'response.end
	    conn.execute(sql)
	
		rs.movenext

	loop
	
	rs.close
	set rs=nothing
	
	response.write("修复进仓结余数据成功!<BR><BR>")








	InstoreRS.close
	set instorers=nothing
	
	set outstoreRS=nothing
	
	conn.close
	set conn=nothing

	Server.ScriptTimeOut =ltimeout
	
	info="数据修复成功!2秒后系统自动返回。"
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