<!-- #include file="data\config.asp" -->
<HTML>
<HEAD>
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<title>���ֵ�ɾ�����һ����</title>
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
 
     Set rs =server.createobject("adodb.recordset") 
     
         
     if request("na")="" then
	     sql="select billno,date1,contractID,goodsname,specs,allnumber,packnumber,weight,ClientName,operater" & " " & _
    	     "from delinstore, ClientInfo" & " " & _
        	 "where ClientInfo.ID=delinstore.ClientID" & " " & _
	         "order by billno asc"
	 else
         if session("ord")<>"asc" then
             session("ord")="asc"
         else 
             session("ord")="desc"
         end if 

	     sql="select billno,date1,contractID,goodsname,specs,allnumber,packnumber,weight,ClientName,operater" & " " & _
    	     "from delinstore, ClientInfo" & " " & _
        	 "where ClientInfo.ID=delinstore.ClientID" & " " & _
	 	     "order by " & request("na") & " " & session("ord")

     end if

 	      

     rs.open sql,conn,1,1
     
     response.write "<table width='900'><tr><td valign='middle' align='center'><b><font size='3'>" & "���ֵ�ɾ�����һ����</font></b></td></tr></table><br>" 
    
     response.write "<table border='1' width='900' bordercolor='#000000'style='border-collapse: collapse' cellspacing='0' cellpadding='0'>"
    
     response.write "<tr><td valign='middle' align='center'><a href='DeleteInstore.asp?na=billno'>����</a></td>" & _
					"<td valign='middle' align='center'><a href='DeleteInstore.asp?na=ClientName'>�������</a></td>" & _
					"<td valign='middle' align='center'><a href='DeleteInstore.asp?na=contractID'>��ͬ��</a></td>" & _
					"<td valign='middle' align='center'><a href='DeleteInstore.asp?na=goodsname'>��Ʒ��</a></td>" & _
					"<td valign='middle' align='center'><a href='DeleteInstore.asp?na=specs'>���</a></td>" & _
					"<td valign='middle' align='center'><a href='DeleteInstore.asp?na=allnumber'>����</a></td>" & _
					"<td valign='middle' align='center'><a href='DeleteInstore.asp?na=packnumber'>����</a></td>" & _
					"<td valign='middle' align='center'><a href='DeleteInstore.asp?na=weight'>����</a></td>" & _
					"<td valign='middle' align='center'><a href='DeleteInstore.asp?na=date1'>����</a></td>" & _
					"<td valign='middle' align='center'>����Ա</td></tr>"
                    
     do while not rs.eof        
        
        response.write "<tr><td valign='middle' align='center'>&nbsp" & rs("billno")  & "&nbsp</td>" & _
                    "<td valign='middle' align='center'>" & left(rs("clientname"),6) & "</td>" & _
                    "<td valign='middle' align='center'>" & rs("contractID") & "</td>" & _
                    "<td valign='middle' align='center'>" & rs("goodsname") & "</td>" & _
                    "<td valign='middle' align='center'>" & rs("specs") & "</td>" & _
                    "<td valign='middle' align='right'>" & rs("allnumber") & "</td>" & _
					"<td valign='middle' align='right'>" & rs("packnumber") & "</td>" & _
                    "<td valign='middle' align='right'>" & gs(cstr(rs("weight")),4) & "</td>" & _
                    "<td valign='middle' align='center'>" & rs("date1") & "</td>" & _
                    "<td valign='middle' align='center'>" & rs("operater") & "</td></tr>"
       
        rs.movenext  
     loop
 
    rs.close
    set rs=nothing
    
    conn.close
	set conn=nothing

%>