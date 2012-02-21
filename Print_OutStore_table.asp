<!-- #include file="data\config.asp" -->


<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>出库通知单</title>
</head>

<body>
<%
	BillNo=Request("BillNo")
	PrintDate=Request("PrintDate")	
	ClientID=Request("ClientID")
	
	Operater1=Session("UserName")
	
	xianshicangwei="OFF"
	
	'''''''''''''''''''''''''''''''''''''''add
	
	XIANSHI=Request("XIANSHI")
	if XIANSHI="ON" then
	   kill_mao="1"
    else
	   kill_mao="0"
    end if
	
	
	
	
	
	'''''''''''''''''''''''''''''''''''''''end
	
	
	
	
	
	

	set rs=server.createobject("adodb.recordset")

	'取公司名称
	sql ="select * from Master"
	rs.open sql,conn,1,1
	if rs.recordcount<1 then
		response.write("本公司信息未录入，请先输入公司信息不存在!")
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing	
		response.redirect "jiliang_list.htm"
		response.end
	end if
	

%>
<table border="0" width="706" cellspacing="0" height="2" cellpadding="0">
  <tr>
    <td colspan="6" width="696" height="46">
<p align="center"><b><font face="宋体" size="5"><% = rs("name") %> </font></b></p>
    </td>
  </tr>
  <tr>
    <td width="690" height="10" colspan="6">

    <p align="center"><b><font face="宋体" size="5"><u>出库通知单</u></font></b>

    </td>
  </tr>
  <tr>
    <td width="77" height="6">

    </td>
    <td width="105" height="6">
    </td>
    <td width="90" height="6">
    </td>
    <td width="166" height="6">
    </td>
    <td width="49" height="6">
      <p align="right">No：</td>
    <td width="203" height="6"><% = BillNo %></td>
  </tr>
  <tr>
    <td width="77" height="9">

    <p align="right"><font size="2">地　址：</font>

    </td>
    <td width="613" height="9" colspan="5"><font size="2"><%=rs("Address") %>
      </font>
    </td>
  </tr>
  <tr>
    <td width="77" height="12">

    <p align="right"><font size="2">电　话：</font>

    </td>
    <td width="613" height="12" colspan="5"><font size="2"><%=rs("Tel")%>&nbsp;&nbsp;&nbsp;&nbsp;传真：<%=rs("Fax")%>
      </font>
    </td>
  </tr>
  <tr>
    <td width="690" height="1" colspan="6" align="center">

    <hr color="#808080" size="1">

    </td>
  </tr>
  </table>

<%
	rs.close
	
	'获得货主公司名称
	sql ="select * from ClientInfo where ID=" & ClientID
	rs.open sql,conn,1,1
	if rs.recordcount<1 then
		response.write("该货主公司不存在!")
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing	
		response.end
	end if
	clientname1=rs("ClientName")
	rs.close   '关闭ClientInfo
	
	sql ="select * from OutStore where BillNo='" & BillNo & "'"
	rs.open sql,conn,1,1
	if rs.recordcount<1 then
		response.write("无出库记录！")
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing	
		response.end
	end if



%>

<table border="0" width="707" cellspacing="0" height="31" cellpadding="0">
  <tr>
    <td width="78" height="20">
<p align="right"><font size="2">起运地：</font>

    </td>
    <td width="294" height="20"><font size="2"><% = rs("StartAddress") %></font></td>
    <td width="80" colspan="2" height="20">
      <p align="right"><font face="宋体" size="2">存货货主：</font></td>
    <td width="239" height="20"><font face="宋体" size="2"><% =ClientName1 %></font></td>
  </tr>
  <tr>
    <td width="78" height="21">	

<p align="right">	

<font face="宋体" size="2">提货单位：</font>
</p>
    </td>
    <td width="613" colspan="4" height="21"><font face="宋体" size="2"><%=rs("Crop")%></font></td>
  </tr>
  <tr>
    <td width="78" height="20">
<p align="right"><font face="宋体" size="2">提货车号：</font>           
    </td>
    <td width="613" colspan="4" height="20"><font size="2"><% =rs("NumberPlate") %></font></td> 
  </tr>
  <tr>
    <td width="78" height="24">
<p align="right">
<font face="宋体" size="2">起运日期：</font>          
</p>
    </td>
    <%
    	datearry=split(Printdate,"-")
    %>

    <td width="613" colspan="4" height="24"><font size="2"><font face="宋体"><%=datearry(0)%>年<%=datearry(1)%>月<%=datearry(2)%>日&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;          
      ***仓库存根(白) 业务部登记(黄) 收货方清单(红) 回单(绿)</font></font></td>             
  </tr>
</table>
<table border="1"  cellspacing="0" bordercolor="#000000" height="29" width="706">
  <tr>
    <td width="90" rowspan="2" height="35" align="center">
    <p align="center"><font size="2">品    名</font></td>                              
    <td width="137" rowspan="2" height="35" align="center"><font size="2">合 同 号</font></td>  
    <td width="102" rowspan="2" height="35" align="center"><font size="2">规 格</font></td>  
    <td width="83" colspan="2" height="15" align="center"><font size="2">数&nbsp;&nbsp; 量</font></td>                            
    <td width="145" colspan="2" height="15" align="center"><font size="2">重 量(吨)</font></td>              
    <td width="123" height="31" rowspan="2" align="center"><font size="2">备 注</font></td>              
  </tr>
  <tr>
    <td width="40" height="16" align="center"><font size="2">捆 数</font></td> 
    <td width="39" height="16" align="center"><font size="2">件 数</font></td>  
    <td width="72" height="16" align="center"><font size="2">毛 重</font></td> 
    <td width="69" height="16" align="center"><font size="2">净 重</font></td>  
  </tr>


<tr>
<%
	NumberSum=0
	PackSum=0
	GrossSum=0
	NetSum=0
	rs.movefirst
	do while not rs.eof
		strGoodsName = trim( rs("GoodsName") )
		if strGoodsName = empty then
			strGoodsName="&nbsp;"
		end if
		strSpecs = trim( rs("Specs") )
		if strSpecs = empty then
			strSpecs = "&nbsp;"
		end if
		strContractID= trim( rs("ContractID") )
		if strContractID = empty then
			strContractID="&nbsp;"
		end if
		strRemark = trim( rs("Remark") )
		
		strPlace2 = trim( rs("Place2") )		
		if strPlace2 = empty then
			strPlace2 = "&nbsp;"
             else
			 
             xianshicangwei= "ON"			
		end if
		
		
		if strRemark = empty then	
		   if xianshicangwei="ON" then
		           strRemark=strPlace2
				   else
			strRemark = "&nbsp;"	
			end if
			
        end if
		
	
		response.write "<td align='center'><font size='2'>" & strGoodsName & "</font></td>" & _
		               "<td align='center'><font size='2'>" & strContractID & "</font></td>" & _
		               "<td align='center'><font size='2'>" & strSpecs & "</font></td>" & _
		               "<td align='right'><font size='2'>" & rs("Number1") & "</font></td>" & _
	    	           "<td align='right'><font size='2'>" & rs("Pack1") & "</font></td>" 
					   
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''		start	屏蔽毛重	   
		 if kill_mao="1" then
          response.write   "<td align='right'><font size='2'>" & gs(cdbl(rs("Gross"))*cdbl(rs("pack1"))/1000,4) & "</font></td>" 
         else
          response.write   "<td align='right'><font size='2'>" & "&nbsp;" & "</font></td>" 
         end if
					   
					   
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''end				   
				   
					   
	  response.write   "<td align='right'><font size='2'>" & gs(cdbl(rs("weight")),4) & "</font></td>" & _
	    	           "<td align='center'>" & strRemark & "</td></tr>"
	    	           
		NumberSum=NumberSum + rs("Number1")
		PackSum=PackSum + rs("Pack1")
		GrossSum=GrossSum + cdbl(rs("Gross"))*cdbl(rs("pack1"))/1000
		NetSum=NetSum + cdbl(rs("weight"))
		Place1=rs("Place1")
		Consigner1=rs("consigner")		
		operater2=rs("operater")

		rs.movenext
	loop	
	
	
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	

%>

  <tr>
    <td width="90" height="25" align="center"><font size="2">&nbsp</font></td>            
    <td width="137" height="25" border="1" align="center"><font size="2">&nbsp</font></td>
    <td width="102" height="25" border="1" align="center"><font size="2">&nbsp</font></td> 
    <td width="40" height="6" border="1" align="right"><font size="2">&nbsp</font></td>
    <td width="39" height="6" border="1" align="right"><font size="2">&nbsp</font></td>
    <td width="72" height="6" border="1" align="right"><font size="2">&nbsp</font></td>
    <td width="69" height="6" border="1" align="right"><font size="2">&nbsp</font></td>
    <td width="123" height="21" border="1" align="center"><font size="2">&nbsp</font></td>
  </tr>

  <tr>
    <td width="90" height="25" align="center"><font size="2">合 计</font></td>                     
    <td width="137" height="25" border="1" align="center"><font size="2">&nbsp</font></td>
    <td width="102" height="25" border="1" align="center"><font size="2">&nbsp</font></td> 
    <td width="40" height="6" border="1" align="right"><font size="2"><%=NumberSum%></font></td>
    <td width="39" height="6" border="1" align="right"><font size="2"><%=PackSum%></font></td>
	
	
	

<%	
''''''''''''''''''''''''''''''''start 屏蔽毛重累计
if kill_mao="1" then
  response.write  "<td width='72' height='6' border='1' align='right'><font size='2'>" & gs(GrossSum,4) & "</font></td>"
else
  response.write  "<td width='72' height='6' border='1' align='right'><font size='2'>"&"&nbsp;"&"</font></td>"
end if
'''''''''''''''''''''''''''''''''end
%>
	
	
	
	<td width="69" height="6" border="1" align="right"><font size="2"><%=gs(NetSum,4)%></font></td>
    <td width="123" height="21" border="1" align="center"><font size="2">&nbsp</font></td>
  </tr>


</table>
<div align="left">
<table border="0" width="707" height="40" cellspacing="1" >
  <tr>
    <td width="69" height="8" valign="middle">
      <p align="right"><font size="2">仓 位：</font>         
    </td>
    <td width="137" height="8" valign="middle">
      <p align="left"><font size="2"><% =place1 %></font></td>
    <td width="66" height="8" valign="middle">
      <p align="right"><font size="2">盖　章：</font></td>
    <td width="95" height="8" valign="middle" align="left">   
    </td>
    <td width="66" height="8" valign="middle">
      <p align="right"><font size="2">复核人：</font></td>
    <td width="68" height="8" valign="middle"></td>
    <td width="86" height="8" valign="middle">   
<p align="right"><font size="2">收货人：</font></p>
    </td>
    <td width="120" height="8" valign="middle"></td>
  </tr>
  <tr>
    <td width="69" height="1" valign="middle"><p align="right"><font size="2">制　单：</font></p>  
    </td>
    <td width="137" height="1" valign="middle">
      <p align="left"><font size="2"><% =operater2 %></font></td>
    <td width="66" height="1" valign="middle">
      <p align="right"><font size="2">发货人：</font></p>
    </td>
    <td width="95" height="1" valign="middle" align="left"><font size="2"><% =consigner1 %></font></td>
    <td width="66" height="1" valign="middle"></td>
    <td width="68" height="1" valign="middle">
      <p align="right"></td>
    <td width="86" height="1" valign="middle">
      <p align="right"><font size="2">身份证号码：</font></td>
    <td width="120" height="1" valign="middle"></td>
  </tr>
</table>

</div>

</body>

</html>
