<!-- #include file="data\config.asp" -->
<html><title>�����޸�</title>
<%
	
	response.write("<BR><BR><BR><center>�˹�����ͣ!<BR><BR>")
	response.write("<a href='jiliang_list.htm'>����</a></center>")
	response.end
    Set rs =server.createobject("adodb.recordset")
	
	ltimeout=Server.ScriptTimeOut 
	Server.ScriptTimeOut =50000
	'���ֽ������ݻָ�Ϊ��ֵ
	%>
	
	
	<script language="vbscript">
		dim choice	
		Choice = msgbox ("��ʼ���������޸���?",vbYESNO,"��ʾ")
		if Choice = vbNO then
			history.back()
		else
			window.alert "׼����ʼ�����ֽ�������!"
		end if
	</script>
	
	<%

	sql = "Update InStore " & _
	      "set remainnumber = allnumber," & _
	      "remainpack = packnumber," & _
	      "remainweight=weight"
	
	conn.Execute(sql)
	
	response.write("��ʼ�����ֽ������ݳɹ�!<BR><BR>")

	

	'ɾ���������ݿ��еĲ����¼

	%>
	<script language="vbscript">
		window.alert "׼���޸�������������!"
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
	response.write("�޸������������ݳɹ�!<BR><BR>")


	'�ý������ݿ��е����ݸ��³��ֿ������

	%>
	<script language="vbscript">
		window.alert "׼�����³�������!"
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
		
		'����δ������,����=����*�ܼ���
		sql = "Update OutStore " & _     
	      	  "set weight=pack1*" & cstr( instoreRS("packweight") ) & "/1000  " & _ 
	      	  "where instoreRecno=" & instoreRS("RecNo") & " and lockweight=FALSE"

		conn.execute(sql)				

		instoreRS.movenext
	
	loop

	response.write("���³������ݳɹ�!<BR><BR>")



	'���ݳ�������,�޸����ֽ�������
	sql = "select * from OutStore"
	rs.open sql,conn,1,1
	
	%>
	<script language="vbscript">
		window.alert "׼���޸����ֽ�������!"
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
	
	response.write("�޸����ֽ������ݳɹ�!<BR><BR>")








	InstoreRS.close
	set instorers=nothing
	
	set outstoreRS=nothing
	
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