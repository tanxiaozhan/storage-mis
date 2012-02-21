<!-- #include file="data\config.asp" -->
<HTML>
<HEAD>
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<title>货物进出情况一览表</title>
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
 	response.write("请选择公司.")
 	response.end
 end if
 
 tmpxiechefei=0            '缷车费临时累计
 allxiechefei=0
 allzhuangchefei=0
 tmpzhuangchefei=0
 
 if trim(request("gsid"))<>"" and request("jy")="计算装卸车费用" then
 
 	bdate= trim( request("tbyear") ) & "-" & trim( request("tbmonth") ) & "-" & trim( request("tbday") )
 	edate= trim( request("teyear") ) & "-" & trim( request("temonth") ) & "-" & trim( request("teday") )
 
     Set rs =server.createobject("adodb.recordset") 
          Set rs10 =server.createobject("adodb.recordset") 
	 
     if trim(request("gsid")) <> "All" then
       	 sql="select ClientName from ClientInfo where ID=" & trim(request("gsid"))
	     rs.open sql,conn,3,3
	     tmpname=rs("ClientName")
         rs.close
     else
     	tmpname="所有客户"
     end if

    for i=0 to 1000
		set rsSchema=Conn.openSchema(20) 
		rsSchema.movefirst 
	    tablename="temp" & trim(cstr(i))
    	x=0 
	    Do Until rsSchema.EOF     '判断表tempX是否存在  X=1---100
			if rsSchema("TABLE_TYPE")="TABLE" then 
				if rsSchema("TABLE_NAME")=tablename then
					x=1 
					exit do
				  end if
    		end if 
			rsSchema.movenext 
		Loop
		
		if x=0 then
		     
		     '创建新表tempX，从进仓，出仓表添加记录
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
			 			 
			 sql="update " & tablename & " set pack='入' "
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
			 sql="update " & tablename & " set pack='出' where pack<>'入' "
			 conn.execute(sql)
			 
			exit for
		end if
	next
     
    '筛选字段
    filte=""
    if request("d4") <> "--" and trim(request("t1")) <> "" then    '筛选1
    		filte="trim(" & request("d3") & ")" & request("d4") & "'" & trim(request("t1")) & "'"
    end if

    if request("d5") <> "--" then
	
	    if request("d6") <> "--" and trim(request("t2")) <> "" then    '筛选2
    				filte=filte & " " & request("d5") & " " & "trim(" & request("d6") & ")" & request("d7") & "'" & trim(request("t2")) & "'"
	    end if
	
	end if
	
  if filte<> "" then
  		filte = " where " & filte
  end if
  
     if request("na")="" then      '排序
	 	sql = "select * from " & tablename &  " " & filte & " order by " & request("d1") & " asc, " & request("d2") & " asc"	
    	sql = "select * from " & tablename & " where pack='入' "
		sql10 = "select * from " & tablename & " where pack<>'入' "
   else
         if session("ord")<>"asc" then
             session("ord")="asc"
         else 
             session("ord")="desc"
         end if 
	' 	sql = "select * from " & tablename & " " & filte & " order by " & request("d1") & " asc, " & request("d2") & " asc, "	& request("na") & " " & session("ord")
  	 	sql = "select * from " & tablename & " where pack='入' "
		sql10 = "select * from " & tablename & " where pack<>'入' "

  end if

   
    rs.open sql,conn,1,1
	
	            if trim(request("D8"))="andunsuan" then   '用于改字段
			     
					xiechefeiziduan="缷车费("&request("xiecheperton")&"元/吨)"   ' 用于调整缷车费字段
   		    else
			     
					xiechefeiziduan="缷车费("&request("xiecheperton")&"元/件)"    ' 用于调整缷车费字段
			end if     

     
     response.write "<table width='900'><tr><td valign='middle' align='center'><b><font size='3'>(" & tmpname & ")卸车费</font></b></td></tr></table><br>" 
    
     response.write "<table border='1' width='900' bordercolor='#000000'style='border-collapse: collapse' cellspacing='0' cellpadding='0'>"
    
'     response.write "<tr><td valign='middle' align='center'><a href='IOBrowseTable.asp?gsid=" & request("gsid") & "&jy=查看进出情况一览表&na=date1'>日期</a></td>" & _
'                    "<td valign='middle' align='center'>进出</a></td>" & _
'     				"<td valign='middle' align='center'><a href='IOBrowseTable.asp?gsid=" & request("gsid") & "&jy=查看进出情况一览表&na=billno'>单  号</a></td>" & _
'     				"<td valign='middle' align='center'><a href='IOBrowseTable.asp?gsid=" & request("gsid") & "&jy=查看进出情况一览表&na=deliver'>提货单位</a></td>" & _
'     				"<td valign='middle' align='center'><a href='IOBrowseTable.asp?gsid=" & request("gsid") & "&jy=查看进出情况一览表&na=ContractID'>合同号</a></td>" & _
'     				"<td valign='middle' align='center'><a href='IOBrowseTable.asp?gsid=" & request("gsid") & "&jy=查看进出情况一览表&na=GoodsName'>产品名</a></td>" & _
'                    "<td valign='middle' align='center'><a href='IOBrowseTable.asp?gsid=" & request("gsid") & "&jy=查看进出情况一览表&na=Specs'>规格</a></td>" & _
'                    "<td valign='middle' align='center'>捆数/小件</a></td>" & _
'                    "<td valign='middle' align='center'><a href='IOBrowseTable.asp?gsid=" & request("gsid") & "&jy=查看进出情况一览表&na=Weight'>净重(吨)</a></td>" & _
'                    "<td valign='middle' align='center'>备注</td></tr>"

     response.write "<tr><td valign='middle' align='center'>日期</a></td>" & _
                    "<td valign='middle' align='center'>进出</a></td>" & _
     				"<td valign='middle' align='center'>单  号</a></td>" & _
     				"<td valign='middle' align='center'>提货单位</a></td>" & _
     				"<td valign='middle' align='center'>合同号</a></td>" & _
     				"<td valign='middle' align='center'>产品名</a></td>" & _
                    "<td valign='middle' align='center'>规格</a></td>" & _
                    "<td valign='middle' align='center'>捆数</a></td>" & _
                    "<td valign='middle' align='center'>小件</a></td>" & _
                    "<td valign='middle' align='center'>净重(吨)</a></td>" & _
                    "<td valign='middle'  align='center'>"&xiechefeiziduan&"</td></tr>"
                    
     do while not rs.eof   

		
            if trim(request("D8"))="andunsuan" then   '用于判断是按吨还是按件数算缷车费
			        tmpxiechefei=gs(cstr(rs("weight")),4)*request("xiecheperton")
   		    else
			        tmpxiechefei=trim(cstr(rs("allnumber")))*request("xiecheperton")
			end if     

			
	  response.write "<tr><td valign='middle' align='center'>" & rs("date1") & "</td>" & _
                    "<td valign='middle' align='center'>" & rs("pack") & "</td>" & _
                    "<td valign='middle' align='center'>" & rs("billno") & "</td>" & _
                    "<td valign='middle' align='left'>" & rs("deliver") & "</td>" & _
                    "<td valign='middle' align='left'>" & rs("ContractID") & "</td>" & _
                    "<td valign='middle' align='left'>" & rs("GoodsName") & "</td>" & _
					"<td valign='middle' align='left'>" & rs("Specs") & "</td>" & _
					"<td valign='middle' align='right'>" & trim(cstr(rs("allnumber"))) & "</td>" & _
					"<td valign='middle' align='right'>" & trim(cstr(rs("packnumber"))) & "</td>" & _
                    "<td valign='middle' align='right'>" & gs(cstr(rs("weight")),4) & "</td>" & _
		
				  "<td valign='middle'  align='center'>￥" & tmpxiechefei & "</td></tr>"
       
   	          allxiechefei=tmpxiechefei+allxiechefei            '缷车费累计
  
  
  
	 
	 rs.movenext  
		
     loop
 
    rs.close
    set rs=nothing
    
	  response.write "<tr><td valign='middle' align='center'>合计</a></td>" & _
                    "<td valign='middle' align='center' colspan='6'></a></td>" & _
					  "<td valign='middle' align='center' colspan='1'></a></td>" & _
					    "<td valign='middle' align='center' colspan='1'></a></td>" & _
						  "<td valign='middle' align='center' colspan='1'></a></td>" & _

                    "<td valign='middle' align='center'>"&round(allxiechefei,2)&"</td></tr>"
       
	
	response.write "  </table>"
	'--------------------------------------------------start 算装车费部分
	rs10.open sql10,conn,1,1

     
	 
	 	        if trim(request("D8"))="andunsuan" then   '改字段
			  
					zhuangchefeiziduan="装车费("&request("zhuangcheperton")&"元/吨)"
   		    else
			 '      
					zhuangchefeiziduan="装车费("&request("zhuangcheperton")&"元/件)"
			end if  
	 
	 response.write "<br>"
	 response.write "<br>"
	 	 response.write "<br>"
     response.write "<table width='900'><tr><td valign='middle' align='center'><b><font size='3'>(" & tmpname & ")装车费</font></b></td></tr></table><br>" 
    
     response.write "<table border='1' width='900' bordercolor='#000000'style='border-collapse: collapse' cellspacing='0' cellpadding='0'>"
    
'     response.write "<tr><td valign='middle' align='center'><a href='IOBrowseTable.asp?gsid=" & request("gsid") & "&jy=查看进出情况一览表&na=date1'>日期</a></td>" & _
'                    "<td valign='middle' align='center'>进出</a></td>" & _
'     				"<td valign='middle' align='center'><a href='IOBrowseTable.asp?gsid=" & request("gsid") & "&jy=查看进出情况一览表&na=billno'>单  号</a></td>" & _
'     				"<td valign='middle' align='center'><a href='IOBrowseTable.asp?gsid=" & request("gsid") & "&jy=查看进出情况一览表&na=deliver'>提货单位</a></td>" & _
'     				"<td valign='middle' align='center'><a href='IOBrowseTable.asp?gsid=" & request("gsid") & "&jy=查看进出情况一览表&na=ContractID'>合同号</a></td>" & _
'     				"<td valign='middle' align='center'><a href='IOBrowseTable.asp?gsid=" & request("gsid") & "&jy=查看进出情况一览表&na=GoodsName'>产品名</a></td>" & _
'                    "<td valign='middle' align='center'><a href='IOBrowseTable.asp?gsid=" & request("gsid") & "&jy=查看进出情况一览表&na=Specs'>规格</a></td>" & _
'                    "<td valign='middle' align='center'>捆数/小件</a></td>" & _
'                    "<td valign='middle' align='center'><a href='IOBrowseTable.asp?gsid=" & request("gsid") & "&jy=查看进出情况一览表&na=Weight'>净重(吨)</a></td>" & _
'                    "<td valign='middle' align='center'>备注</td></tr>"

     response.write "<tr><td valign='middle' align='center'>日期</a></td>" 
     response.write               "<td valign='middle' align='center'>进出</a></td>" 
     response.write				"<td valign='middle' align='center'>单  号</a></td>" 
     response.write				"<td valign='middle' align='center'>提货单位</a></td>" 
     response.write				"<td valign='middle' align='center'>合同号</a></td>" 
     response.write				"<td valign='middle' align='center'>产品名</a></td>" 
     response.write               "<td valign='middle' align='center'>规格</a></td>" 
     response.write               "<td valign='middle' align='center'>捆数</a></td>" 
     response.write               "<td valign='middle' align='center'>小件</a></td>" 
     response.write               "<td valign='middle' align='center'>净重(吨)</a></td>" 
     response.write               "<td valign='middle' align='center'>"&zhuangchefeiziduan&"</td></tr>"
          

          
     do while not rs10.eof        
      
	  
	        if trim(request("D8"))="andunsuan" then   '用于判断是按吨还是按件数算装车费
			        tmpzhuangchefei=gs(cstr(rs10("weight")),4)*request("zhuangcheperton")
   		    else
			        tmpzhuangchefei=trim(cstr(rs10("allnumber")))*request("zhuangcheperton")
			end if   
	  
	  
        response.write "<tr><td valign='middle' align='center'>" & rs10("date1") & "</td>" & _
                    "<td valign='middle' align='center'>" & rs10("pack") & "</td>" & _
                    "<td valign='middle' align='center'>" & rs10("billno") & "</td>" & _
                    "<td valign='middle' align='left'>" & rs10("deliver") & "</td>" & _
                    "<td valign='middle' align='left'>" & rs10("ContractID") & "</td>" & _
                    "<td valign='middle' align='left'>" & rs10("GoodsName") & "</td>" & _
					"<td valign='middle' align='left'>" & rs10("Specs") & "</td>" & _
					"<td valign='middle' align='right'>" & trim(cstr(rs10("allnumber"))) & "</td>" & _
					"<td valign='middle' align='right'>" & trim(cstr(rs10("packnumber"))) & "</td>" & _
                    "<td valign='middle' align='right'>" & gs(cstr(rs10("weight")),4) & "</td>" & _
                    "<td valign='middle' align='center'>￥" & tmpzhuangchefei& "</td></tr>"
       
	     allzhuangchefei=tmpzhuangchefei+allzhuangchefei 
        rs10.movenext  
		
     loop
 
    rs10.close
    set rs10=nothing
	
	   response.write "<tr><td valign='middle' align='center'>合计</a></td>" & _
                    "<td valign='middle' align='center' colspan='6'></a></td>" & _
					 "<td valign='middle' align='center' colspan='1'></a></td>" & _
					  "<td valign='middle' align='center' colspan='1'></a></td>" & _
					   "<td valign='middle' align='center' colspan='1'></a></td>" & _
					
					
     				  "<td valign='middle' align='center'>"&round(allzhuangchefei,2)&"</td></tr>"
       
	
	response.write "</table>"
	
	'----------------------------------------------------------end 算装车费部分结束
    '删除表
    sql="drop table " & tablename 
    conn.execute(sql)
    
    conn.close
	set conn=nothing
end if

%>

