<!-- #include file="data\config.asp" -->
<HTML>
<HEAD>
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<title>����������һ����</title>
<STYLE type=text/css>
<!--
A:link {text-decoration:none; color:"#0000ff"}
A:visited {text-decoration:none}
A:hover {text-decoration:underline; color:"#FF0033"}

body { font-family:"Verdana", "Arial", "Helvetica", "sans-serif"; font-size: 9pt}

p {  font-family:"Verdana", "Arial", "Helvetica", "sans-serif"; font-size: 9pt}

td {  font-family: "Verdana", "Arial", "Helvetica", "sans-serif"; font-size: 9pt}
-->
</STYLE>
</HEAD>
<%
 if trim(request("gsid"))="" then
 	response.write("��ѡ��˾.")
 	response.end
 end if
 
 
 if trim(request("gsid"))<>"" and request("jy")="�鿴�������һ����" then
 
 	bdate= trim( request("tbyear") ) & "-" & trim( request("tbmonth") ) & "-" & trim( request("tbday") )
 	edate= trim( request("teyear") ) & "-" & trim( request("temonth") ) & "-" & trim( request("teday") )
 
     Set rs =server.createobject("adodb.recordset") 
     
     if trim(request("gsid")) <> "All" then
       	 sql="select ClientName from ClientInfo where ID=" & trim(request("gsid"))
	     rs.open sql,conn,3,3
	     tmpname=rs("ClientName")
         rs.close
     else
     	tmpname="���пͻ�"
     end if

    for i=0 to 1000
		set rsSchema=Conn.openSchema(20) 
		rsSchema.movefirst 
	    tablename="temp" & trim(cstr(i))
    	x=0 
	    Do Until rsSchema.EOF     '�жϱ�tempX�Ƿ����  X=1---100
			if rsSchema("TABLE_TYPE")="TABLE" then 
				if rsSchema("TABLE_NAME")=tablename then
					x=1 
					exit do
				  end if
    		end if 
			rsSchema.movenext 
		Loop
		
		if x=0 then
		     
		     '�����±�tempX���ӽ��֣����ֱ���Ӽ�¼
		     if trim(request("gsid")) = "All" then
		          sql="select date1,pack,billno,Deliver,contractID,goodsname,specs,allnumber,packnumber,weight,remark,time1 into " & tablename & " " & _
		     	    "from instore " & _
		            "where date1>=#" & bdate & "# and date1<=#" & edate & "#"
		     
		     else
		          sql="select date1,pack,billno,Deliver,contractID,goodsname,specs,allnumber,packnumber,weight,remark,time1 into " & tablename & " " & _
		     	    "from instore " & _
		            "where ClientID=" & request("gsid") & " and date1>=#" & bdate & "# and date1<=#" & edate & "#"
		     end if
			 
			 conn.Execute(sql)
			 			 
			 sql="update " & tablename & " set pack='��' "
			 conn.execute(sql)
			 
			 if trim(request("gsid"))="All" then
			 	 sql="insert into " & tablename & " select date2 as date1,pack,billno,crop as deliver,contractID,goodsname,specs,number1 as allnumber,pack1 as packnumber,weight,remark,time1 " & _
			         "from outstore " & _
		             "where date2>=#" & bdate & "# and date2<=#" & edate & "#"
		     else
			 	 sql="insert into " & tablename & " select date2 as date1,pack,billno,crop as deliver,contractID,goodsname,specs,number1 as allnumber,pack1 as packnumber,weight,remark,time1 " & _
			         "from outstore " & _
		             "where ClientID=" & request("gsid") & "and date2>=#" & bdate & "# and date2<=#" & edate & "#"
		     end if
		         
			 conn.execute(sql)
			 sql="update " & tablename & " set pack='��' where pack<>'��' "
			 conn.execute(sql)
			 
			exit for
		end if
	next
     
    'ɸѡ�ֶ�
    filte=""
    if request("d4") <> "--" and trim(request("t1")) <> "" then    'ɸѡ1
    		filte="trim(" & request("d3") & ")" & request("d4") & "'" & trim(request("t1")) & "'"
    end if

    if request("d5") <> "--" then
	
	    if request("d6") <> "--" and trim(request("t2")) <> "" then    'ɸѡ2
    				filte=filte & " " & request("d5") & " " & "trim(" & request("d6") & ")" & request("d7") & "'" & trim(request("t2")) & "'"
	    end if
	
	end if
	
  if filte<> "" then
  		filte = " where " & filte
  end if
  
     if request("na")="" then      '����
	 	sql = "select * from " & tablename &  " " & filte & " order by " & request("d1") & " asc, " & request("d2") & " asc"	
     else
         if session("ord")<>"asc" then
             session("ord")="asc"
         else 
             session("ord")="desc"
         end if 
	 	sql = "select * from " & tablename & " " & filte & " order by " & request("d1") & " asc, " & request("d2") & " asc, "	& request("na") & " " & session("ord")
     end if


    rs.open sql,conn,1,1
     
     response.write "<table width='900'><tr><td valign='middle' align='center'><b><font size='3'>(" & tmpname & ")�������һ����</font></b></td></tr></table><br>" 
    
     response.write "<table border='1' width='900' bordercolor='#000000'style='border-collapse: collapse' cellspacing='0' cellpadding='0'>"
    
'     response.write "<tr><td valign='middle' align='center'><a href='IOBrowseTable.asp?gsid=" & request("gsid") & "&jy=�鿴�������һ����&na=date1'>����</a></td>" & _
'                    "<td valign='middle' align='center'>����</a></td>" & _
'     				"<td valign='middle' align='center'><a href='IOBrowseTable.asp?gsid=" & request("gsid") & "&jy=�鿴�������һ����&na=billno'>��  ��</a></td>" & _
'     				"<td valign='middle' align='center'><a href='IOBrowseTable.asp?gsid=" & request("gsid") & "&jy=�鿴�������һ����&na=deliver'>�����λ</a></td>" & _
'     				"<td valign='middle' align='center'><a href='IOBrowseTable.asp?gsid=" & request("gsid") & "&jy=�鿴�������һ����&na=ContractID'>��ͬ��</a></td>" & _
'     				"<td valign='middle' align='center'><a href='IOBrowseTable.asp?gsid=" & request("gsid") & "&jy=�鿴�������һ����&na=GoodsName'>��Ʒ��</a></td>" & _
'                    "<td valign='middle' align='center'><a href='IOBrowseTable.asp?gsid=" & request("gsid") & "&jy=�鿴�������һ����&na=Specs'>���</a></td>" & _
'                    "<td valign='middle' align='center'>����/С��</a></td>" & _
'                    "<td valign='middle' align='center'><a href='IOBrowseTable.asp?gsid=" & request("gsid") & "&jy=�鿴�������һ����&na=Weight'>����(��)</a></td>" & _
'                    "<td valign='middle' align='center'>��ע</td></tr>"

     response.write "<tr><td valign='middle' align='center'>����</a></td>" & _
                    "<td valign='middle' align='center'>����</a></td>" & _
     				"<td valign='middle' align='center'>��  ��</a></td>" & _
     				"<td valign='middle' align='center'>�����λ</a></td>" & _
     				"<td valign='middle' align='center'>��ͬ��</a></td>" & _
     				"<td valign='middle' align='center'>��Ʒ��</a></td>" & _
                    "<td valign='middle' align='center'>���</a></td>" & _
                    "<td valign='middle' align='center'>����</a></td>" & _
                    "<td valign='middle' align='center'>С��</a></td>" & _
                    "<td valign='middle' align='center'>����(��)</a></td>" & _
                    "<td valign='middle' align='center'>��ע</td></tr>"
                    
     do while not rs.eof        
        
        response.write "<tr><td valign='middle' align='center'>" & rs("date1") & "</td>" & _
                    "<td valign='middle' align='center'>" & rs("pack") & "</td>" & _
                    "<td valign='middle' align='center'>" & rs("billno") & "</td>" & _
                    "<td valign='middle' align='left'>" & left(rs("deliver"),6) & "</td>" & _
                    "<td valign='middle' align='left'>" & rs("ContractID") & "</td>" & _
                    "<td valign='middle' align='left'>" & rs("GoodsName") & "</td>" & _
					"<td valign='middle' align='left'>" & rs("Specs") & "</td>" & _
					"<td valign='middle' align='right'>" & trim(cstr(rs("allnumber"))) & "</td>" & _
					"<td valign='middle' align='right'>" & trim(cstr(rs("packnumber"))) & "</td>" & _
                    "<td valign='middle' align='right'>" & gs(cstr(rs("weight")),4) & "</td>" & _
                    "<td valign='middle' align='left'>" & rs("remark") & "</td></tr>"
       
        rs.movenext  
     loop
 
    rs.close
    set rs=nothing
    
    'ɾ����
    sql="drop table " & tablename 
    conn.execute(sql)
    
    conn.close
	set conn=nothing
end if

%>

