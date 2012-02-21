<!-- #include file="data\config.asp" -->


<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>入库通知单</title>
</head>

<body>
<%
	BillNo=Request("BillNo")
	PrintDate=Request("PrintDate")	
	ClientID=Request("ClientID")
	
	Operater1=Session("UserName")

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
<table border="0" width="686" cellspacing="0" height="1" cellpadding="0">
  <tr>
    <td colspan="6" width="676" height="46">
<p align="center"><b><font face="宋体" size="5"><% = rs("name") %> </font></b></p>
    </td>
  </tr>
  <tr>
    <td width="670" height="6" colspan="6">

    <p align="center"><b><font face="宋体" size="5"><u>入库通知单</u></font></b>

    </td>
  </tr>
  <tr>
    <td width="92" height="6">

    </td>
    <td width="90" height="6">
    </td>
    <td width="90" height="6">
    </td>
    <td width="166" height="6">
    </td>
    <td width="49" height="6">
      <p align="right">No：</td>
    <td width="183" height="6"><font size="2"><% = BillNo %></font></td>
  </tr>
  <tr>
    <td width="92" height="6">

    <p align="right"><font size="2">地　址：</font>

    </td>
    <td width="578" height="6" colspan="5"><font size="2"><%=rs("Address") %>
      </font>
    </td>
  </tr>
  <tr>
    <td width="92" height="6">

    <p align="right"><font size="2">电　话：</font>

    </td>
    <td width="578" height="6" colspan="5"><font size="2"><%=rs("Tel")%>&nbsp;&nbsp;&nbsp;&nbsp;传真：<%=rs("Fax")%>
      </font>
    </td>
  </tr>
  <tr>
    <td width="670" height="1" colspan="6" align="center">

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
	
	sql ="select * from InStore where BillNo='" & BillNo & "'"
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

<table border="0" width="686" cellspacing="0" height="5" cellpadding="0">
  <tr>
    <td width="92" height="20">
      <p align="right"><font face="宋体" size="2">存货货主：</font>

    </td>
    <td width="290" height="20"><font size="2"><% =ClientName1 %></font></td>
    <td width="121" height="20">
      <p align="right"><font size="2" face="宋体">批号(合同号)：</font></td>
    <td width="167" height="20"><font face="宋体" size="2"><% =rs("ContractID") %></font></td>
  </tr>
  <tr>
    <td width="92" height="20">
<p align="right"><font size="2" face="宋体">送货单位：</font>

    </td>
    <td width="290" height="20"><font size="2"><% =rs("Deliver") %></font></td>
    <td width="121" height="20">
      <p align="right"><font face="宋体" size="2">入库库位：</font></td>
    <td width="167" height="20"><font size="2"><% = rs("Place1") %></font></td>
  </tr>
  <tr>
    <td width="92" height="1">
<p style="line-height: 100%; margin-top: 5; margin-bottom: 5" align="right"><font face="宋体" size="2">入库日期：</font></p>           
    </td>
    <%
    	datearry=split(Printdate,"-")
    %>
    <td width="578" colspan="3" height="6"><font size="2"><font face="宋体"><%=datearry(0)%>年<%=datearry(1)%>月<%=datearry(2)%>日&nbsp;               
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ***仓库存根(白) 业务部登记(黄) 收货方清单(红) 回单(绿)</font></font></td>                    
  </tr>
</table>
<table border="1"  cellspacing="0" bordercolor="#000000" height="29" width="684">
  <tr>
    <td width="127" rowspan="2" height="35" >
    <p align="center"><font size="2">品 名</font></td>                                 
    <td width="128" rowspan="2" height="35" align="center">
      <p align="center"><font size="2">规 格</font></p>
    </td>  
    <td width="93" colspan="2" height="15" align="center">
      <p align="center"><font size="2">数 量</font></p> 
    </td>                           
    <td width="170" colspan="2" height="15" align="center"><font size="2">重量(吨)</font></td>              
    <td width="144" height="31" rowspan="2" align="center">
      <p align="center"><font size="2">备 注</font></p>          
    </td>        
  </tr>
  <tr>
    <td width="41" height="16" align="center">
      <p align="center"><font size="2">捆 数</font></p>
    </td> 
    <td width="48" height="16" align="center"><font size="2">件 数</font></td>  
    <td width="83" height="16" align="center"><font size="2">毛 重</font></td> 
    <td width="83" height="16" align="center"><font size="2">净 重</font></td>  
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
		strRemark = trim( rs("Remark") )
		if strRemark = empty then
			strRemark = "&nbsp;"
		end if
		
		response.write "<td align='center'><font size='2'>" & rs("GoodsName") & "</font></td>" & _
		               "<td align='center'><font size='2'>" & rs("Specs") & "</font></td>" & _
		               "<td align='right'><font size='2'>" & rs("AllNumber") & "</font></td>" & _
	    	           "<td align='right'><font size='2'>" & rs("PackNumber") & "</font></td>" & _
	    	           "<td align='right'><font size='2'>" & trim(gs(cdbl(rs("Gross"))*cdbl(rs("PackNumber"))/1000,4)) & "</font></td>" & _
	    	           "<td align='right'><font size='2'>" & trim(gs(cdbl(rs("Weight")),4)) & "</font></td>" & _
	    	           "<td>" & strRemark  & "</td></tr>"
		NumberSum=NumberSum + rs("AllNumber")
		PackSum=PackSum + rs("PackNumber")
		GrossSum=GrossSum + cdbl(rs("Gross")) * cdbl(rs("PackNumber"))/1000
		NetSum=NetSum + cdbl(rs("weight"))
		Consignee1=rs("Consignee")
		operater2=rs("operater")		
		rs.movenext
	loop	
	
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	

%>

  <tr>
    <td width="127" height="25" align="center"><font size="2">&nbsp</font></td>            
    <td width="128" height="25" border="1" align="center"><font size="2">&nbsp</font></td> 
    <td width="41" height="6" border="1" align="right"><font size="2">&nbsp</font></td>
    <td width="48" height="6" border="1" align="right"><font size="2">&nbsp</font></td>
    <td width="83" height="6" border="1" align="right"><font size="2">&nbsp</font></td>
    <td width="83" height="6" border="1" align="right"><font size="2">&nbsp</font></td>
    <td width="144" height="21" border="1" align="center"><font size="2">&nbsp</font></td>
  </tr>

  <tr>
    <td  align="center" width="127"><font size="2">合计</font></td>                     
    <td  align="center" width="128"><font size="2">&nbsp</font></td> 
    <td  align="right" width="41"><font size="2"><%=NumberSum%></font></td>
    <td  align="right" width="48"><font size="2"><%=PackSum%></font></td>
    <td  align="right" width="83"><font size="2"><%=gs(GrossSum,4)%></font></td>
    <td  align="right" width="83"><font size="2"><%=gs(NetSum,4)%></font></td>
    <td  align="center" width="144"><font size="2">&nbsp</font></td>
  </tr>


</table>
<div align="left">
<table border="0" width="684" height="1" cellspacing="1" >
  <tr>
    <td width="89" height="1" valign="middle">
      <p align="right"><font size="2">收货人：</font>  
    </td>
    <td width="166" height="1" valign="middle">
      <p align="left"><font size="2"><% =Consignee1 %></font>
    </td>
    <td width="165" height="1" valign="middle">
      <p align="right"><font size="2">制　单：</font></td>
    <td width="238" height="1" valign="middle" ><font size="2"><% =Operater2 %></font></td>
  </tr>
  <tr>
    <td width="89" height="8" valign="middle">  
    </td>
    <td width="166" height="8" valign="middle">
      <p align="right"></p>
    </td>
    <td width="165" height="8" valign="middle">
      <p align="right"><font size="2" face="宋体">入库签收（盖章）：</font></td>
    <td width="238" height="8" valign="middle"><font size="2"></font>
      <p align="right">
      </td>
  </tr>
</table>

</div>

</body>

</html>
