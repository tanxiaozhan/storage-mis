<!-- #include file="data\config.asp" -->


<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>����֪ͨ��</title>
</head>

<body>
<%
	BillNo=Request("BillNo")
	PrintDate=Request("PrintDate")	
	ClientID=Request("ClientID")
	
	Operater1=Session("UserName")

	set rs=server.createobject("adodb.recordset")

	'ȡ��˾����
	sql ="select * from Master"
	rs.open sql,conn,1,1
	if rs.recordcount<1 then
		response.write("����˾��Ϣδ¼�룬�������빫˾��Ϣ������!")
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing	
		response.redirect "jiliang_list.htm"
		response.end
	end if
	

%>
<table border="0"  cellspacing="1" width="694">
  <tr>
    <td width="686">
<p align="center"><b><font face="����" size="5"><% = rs("name") %> </font></b></p>
    </td>
  </tr>
  <tr>
    <td width="686">
      <p align="center"><font face="����" size="5"><u><b>ǩ��֪ͨ��</b></u></font>

    </td>
  </tr>
</table>
<table border="0"  cellspacing="1" width="696">
  <tr>
    <td width="76"></td>
    <td width="312"></td>
    <td width="90">
      <p align="right">No��</td>
    <td width="192"><font size="2"><% = BillNo %></font></td>
  </tr>
  <tr>
    <td width="76" align="right">
      <p align="right"><font size="2">�ء�ַ��</font>

    </td>
    <td width="606" colspan="3"><font size="2"><%=rs("Address") %></font>
    </td>
  </tr>
  <tr>
    <td width="76" align="right"><font size="2">�硡����</font>

    </td>
    <td width="606" colspan="3"><font size="2"><%=rs("Tel")%>&nbsp;&nbsp;&nbsp;&nbsp;���棺<%=rs("Fax")%>

      </font>
    </td>
  </tr>
  <tr>
    <td width="688" colspan="4">
      <hr>
    </td>
  </tr>
</table>

<%
	rs.close
	
	'��û�����˾����
	sql ="select * from ClientInfo where ID=" & ClientID
	rs.open sql,conn,1,1
	if rs.recordcount<1 then
		response.write("�û�����˾������!")
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing	
		response.end
	end if
	clientname1=rs("ClientName")
	rs.close   '�ر�ClientInfo
	
	sql ="select * from OutStore where BillNo='" & BillNo & "'"
	rs.open sql,conn,1,1
	if rs.recordcount<1 then
		response.write("�޳����¼��")
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing	
		response.end
	end if



%>





<table border="0"  cellspacing="1" width="698">
  <tr>
    <td width="76" align="right"><font size="2" face="����">���˵أ�</font></td>
    <td width="269"><font size="2"><% = rs("StartAddress") %></font></td>
    <td width="114" align="right"><font face="����" size="2">�ջ���λ��</font></td>
    <td width="213"><font size="2">
      <p align="left"><%=rs("Crop")%></font></td>
  </tr>
  <tr>
    <td width="76" align="right"><font size="2" face="����">���������</font></td>
    <td width="269"><font face="����" size="2"><%=ClientName1%></font></td>
    <td width="114" align="right"><font face="����" size="2">�ͻ���ַ��</font></td>
    <td width="213"><font size="2"></font>
      <p align="left"></td>
  </tr>
  <tr>
    <td width="76" align="right"><font size="2" face="����">�ͻ����ţ�</font></td>
    <td width="269"><font size="2"><% =rs("NumberPlate") %></font></td>
    <td width="114" align="right"><font face="����" size="2">��ϵ��&amp;��ϵ�绰��</font></td>
    <td width="213"><font size="2"></font>
      <p align="left"></td>
  </tr>
  <tr>
    <%
    	datearry=split(Printdate,"-")
    %>
    <td width="76" align="right"><font size="2" face="����">�������ڣ�</font></td>
    <td width="608" colspan="3"><font face="����" size="2"><%=datearry(0)%>��<%=datearry(1)%>��<%=datearry(2)%>��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;               
      ***�ֿ���(��) ҵ�񲿵Ǽ�(��) �ջ����嵥(��) �ص�(��)</font></td>            
  </tr>
</table>
<table border="1"  cellspacing="0" bordercolor="#000000" height="29" width="696">
  <tr>
    <td width="120" rowspan="2" height="35" align="center">
    <p align="center"><font size="2">Ʒ ��</font></td>                                  
    <td width="127" rowspan="2" height="35" align="center"><font size="2">�� ��</font></td>  
    <td width="78" colspan="2" height="15" align="center"><font size="2">��&nbsp;&nbsp; ��</font></td>                                  
    <td width="208" colspan="2" height="15" align="center"><font size="2">�� ��(��)</font></td>                   
    <td width="141" height="31" rowspan="2" align="center"><font size="2">�� ע</font></td>                   
  </tr>
  <tr>
    <td width="36" height="16" align="center">
      <p align="center"><font size="2">�� ��</p>
    </td> 
    <td width="38" height="16" align="center">
      <p align="center"><font size="2">�� ��</font></p> 
    </td>  
    <td width="101" height="16" align="center"><font size="2">ë ��</font></td> 
    <td width="103" height="16" align="center"><font size="2">�� ��</font></td>  
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
		               "<td align='right'><font size='2'>" & rs("Number1") & "</font></td>" & _
	    	           "<td align='right'><font size='2'>" & rs("Pack1") & "</font></td>" & _
	    	           "<td align='right'><font size='2'>" & gs(cdbl(rs("Gross"))*cdbl(rs("pack1"))/1000,4) & "</font></td>" & _
	    	           "<td align='right'><font size='2'>" & gs(cdbl(rs("weight")),4) & "</font></td>" & _
	    	           "<td>" & strRemark  & "</td></tr>"
	    	           
		NumberSum=NumberSum + rs("Number1")
		PackSum=PackSum + rs("Pack1")
		GrossSum=GrossSum + cdbl(rs("Gross"))*cdbl(rs("pack1"))/1000
		NetSum=NetSum + cdbl(rs("weight"))
		Consigner1=rs("consigner")		
		Place11=rs("Place1")
		operater2=rs("operater")
		rs.movenext
	loop	
	
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	

%>

  <tr>
    <td width="120" height="25" align="center">&nbsp</td>            
    <td width="127" height="25" border="1" align="center">&nbsp</td> 
    <td width="36" height="6" border="1" align="right">&nbsp</td>
    <td width="38" height="6" border="1" align="right">&nbsp</td>
    <td width="101" height="6" border="1" align="right"><font size="2">&nbsp</font></td>
    <td width="103" height="6" border="1" align="right"><font size="2">&nbsp</font></td>
    <td width="141" height="6" border="1" align="right"><font size="2">&nbsp</font></td>
  </tr>

  <tr>
    <td  align="center" width="120"><font size="2">�� ��</font></td>                          
    <td  align="center" width="127"><font size="2">&nbsp</font></td> 
    <td  align="right" width="36"><font size="2"><%=NumberSum%></font></td>
    <td  align="right" width="38"><font size="2"><%=PackSum%></font></td>
    <td  align="right" width="101"><font size="2"><%=gs(GrossSum,4)%></font></td>
    <td  align="right" width="103"><font size="2"><%=gs(NetSum,4)%></font></td>
    <td  align="right" width="141"><font size="2">&nbsp</font></td>
    
  </tr>


</table>
<div align="left">
<table border="0" width="697" height="40" cellspacing="1" >
  <tr>
    <td width="67" height="8" valign="middle">
      <p align="right"><font size="2">�� λ��</font>              
    </td>
    <td width="138" height="8" valign="middle">
      <p align="left"><font size="2"><%=place11%></font></td>
    <td width="54" height="8" valign="middle">
      <p align="right"><font size="2">�ǡ��£�</font></td>
    <td width="95" height="8" valign="middle" align="left">   
    </td>
    <td width="58" height="8" valign="middle">
      <p align="right"><font size="2">�����ˣ�</font></td>
    <td width="86" height="8" valign="middle"></td>
    <td width="58" height="8" valign="middle">   
<p align="right"><font size="2">�ջ��ˣ�</font></p>
    </td>
    <td width="115" height="8" valign="middle"></td>
  </tr>
  <tr>
    <td width="67" height="1" valign="middle"><p align="right"><font size="2">�ơ�����</font></p>  
    </td>
    <td width="138" height="1" valign="middle">
      <p align="left"><font size="2"><% =Operater2 %></font></td>
    <td width="54" height="1" valign="middle">
      <p align="right"><font size="2">�����ˣ�</font></p>
    </td>
    <td width="95" height="1" valign="middle" align="left"><font size="2"><% =consigner1 %></font></td>
    <td width="58" height="1" valign="middle"></td>
    <td width="86" height="1" valign="middle">
      <p align="right"></td>
    <td width="58" height="1" valign="middle">
      <p align="right"><font size="2">�ǡ��£�</font></td>
    <td width="115" height="1" valign="middle"></td>
  </tr>
</table>

</div>

</body>

</html>
