<!-- #include file="data\config.asp" -->
<HTML>
<HEAD>
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<title>�ܽ�����</title>
<STYLE type=text/css>
<!--
A:link {text-decoration:none; color:"#0000ff"}
A:visited {text-decoration:none}
A:hover {text-decoration:underline; color:"#FF0033"}

body { font-family:"Verdana", "Arial", "Helvetica", "sans-serif"; font-size: 11pt}

p {  font-family:"Verdana", "Arial", "Helvetica", "sans-serif"; font-size: 11pt}

td {  font-family: "Verdana", "Arial", "Helvetica", "sans-serif"; font-size: 11pt}
-->
</STYLE>
</HEAD>
<%
 if request("jy")="�鿴���" and trim(request("ye"))<>"" and trim(request("mon"))<>"" and trim(request("ri"))<>"" then
     
	 if not isdate(trim(request("ye")) & "-" & trim(request("mon")) & "-" & trim(request("ri"))) then
        conn.close
        set conn=nothing
        response.write "���ڳ������磺�е��·���31�գ��е��·���30�գ��е��·���29�գ��е��·���28�գ�"
        response.end  
     end if
	 
	 if cdate(trim(request("ye")) & "-" & trim(request("mon")) & "-" & trim(request("ri"))) > date() then
         response.write "��[ѡ���ʱ��]����[��ǰʱ��]��"	    
		 conn.close
		 set conn=nothing
		 response.end
     end if
	 
	 tmpday1=cdate(trim(request("ye")) & "-" & trim(request("mon")) & "-" & trim(request("ri")))
	 
     Set rs =server.createobject("adodb.recordset") 
	 
	 Set rs1 =server.createobject("adodb.recordset") 
     
     'sql="select ClientName from ClientInfo where ID=" & trim(request("gsid"))
     
     'rs.open sql,conn,3,3
     
     'tmpname=rs("ClientName")

     'rs.close
     
     if request("na")="" then
        'if request("lin")="zero" then
        '  sql="select ContractID,GoodsName,Specs,remainnumber,remainweight,Date1,Time1,remark from InStore where ClientId=" & trim(request("gsid")) & " and remainnumber>0 and remainweight>0 order by Date1 asc,Time1 asc"
        'else  
         sql="select ClientName,RecNo,ContractID,GoodsName,Specs,remainnumber,AllNumber,remainweight,Date1,Time1,remark from ClientInfo,InStore where ClientInfo.ID=InStore.ClientID and date1<=#" & tmpday1 & "# order by ClientName asc,Date1 asc,Time1 asc"
        'end if
     else
         if session("ord")<>"asc" then
             session("ord")="asc"
         else 
             session("ord")="desc"
         end if 

         'if request("lin")="zero" then
         '   sql="select ContractID,GoodsName,Specs,remainnumber,remainweight,Date1,Time1,remark from InStore where ClientId=" & trim(request("gsid")) & " and remainnumber>0 and remainweight>0 order by " & request("na") & " " & session("ord")
         'else  
         sql="select ClientName,RecNo,ContractID,GoodsName,Specs,remainnumber,AllNumber,remainweight,Date1,Time1,remark from ClientInfo,InStore where ClientInfo.ID=InStore.ClientID and date1<=#" & tmpday1 & "# order by ClientName asc," & request("na") & " " & session("ord") & ",Date1 asc,Time1 asc"
         'end if
 
     end if


     rs.open sql,conn,3,3
	 
     response.write "<p>"  
     response.write "<table width='1000'><tr><td valign='middle' align='center'><b><font size='3'>�ܽ�����(" & tmpday1 & ")</font></b></td></tr></table>" 
    
     response.write "<table border='1' width='1000' bordercolor='#000000'style='border-collapse: collapse' cellspacing='0' cellpadding='0'>"

      
     response.write "<tr><td valign='middle' align='center'>���</td><td valign='middle' align='center'>��������</td><td valign='middle' align='center'><a href='cremainty1.asp?jy=�鿴���&na=ContractID&ye=" & trim(request("ye")) & "&mon=" & trim(request("mon")) & "&ri=" & trim(request("ri")) & "'>��ͬ��</a></td><td valign='middle' align='center'><a href='cremainty1.asp?jy=�鿴���&na=GoodsName&ye=" & trim(request("ye")) & "&mon=" & trim(request("mon")) & "&ri=" & trim(request("ri")) & "'>��Ʒ��</a></td><td valign='middle' align='center'><a href='cremainty1.asp?jy=�鿴���&na=Specs&ye=" & trim(request("ye")) & "&mon=" & trim(request("mon")) & "&ri=" & trim(request("ri")) & "'>���</a></td>" & _
                    "<td valign='middle' align='center'><a href='cremainty1.asp?jy=�鿴���&na=remainnumber&ye=" & trim(request("ye")) & "&mon=" & trim(request("mon")) & "&ri=" & trim(request("ri")) & "'>Ŀǰ�������/����</a></td><td valign='middle' align='center'><a href='cremainty1.asp?jy=�鿴���&na=remainweight&ye=" & trim(request("ye")) & "&mon=" & trim(request("mon")) & "&ri=" & trim(request("ri")) & "'>Ŀǰ������</a></td><td valign='middle' align='center'><a href='cremainty1.asp?jy=�鿴���&na=date1&ye=" & trim(request("ye")) & "&mon=" & trim(request("mon")) & "&ri=" & trim(request("ri")) & "'>����ʱ��</a></td><td valign='middle' align='center'>��ע</td></tr>"
          
     ii=0
        
        zremainw=0
        zremainjian=0

        zprices=0
         
		ktv=0
		colsed=41     
     do while not rs.eof        
        ii=ii+1
        
        'if (cdbl(rs("remainweight"))<1) and (cdbl(rs("remainweight"))>0) then 
        '    remainw= "0" & rs("remainweight")
        'else
        'remainw=rs("remainweight")    
        'end if
        sql1="select sum(Number1) as zjs,sum(weight) as zw from OutStore where InStoreRecNo=" & rs("RecNo") & " and Date1>#" & tmpday1 & "#"
		
		rs1.open sql1,conn,3,3
		
		if rs1.recordcount>0 then
		   if not isnull(rs1("zjs")) then
		     zjs1=cdbl(rs1("zjs"))+cdbl(rs("remainnumber"))
		   else
		     zjs1=cdbl(rs("remainnumber")) 	 
		   end if
		   
		   if not isnull(rs1("zw")) then
		     zw1=cdbl(rs1("zw"))+cdbl(rs("remainweight"))
		   else
		     zw1=cdbl(rs("remainweight"))	 
		   end if
		else
		   zjs1=cdbl(rs("remainnumber"))
		   zw1=cdbl(rs("remainweight"))	   
		end if
		
		rs1.close
		
	    if request("lin")="zero" then
              response.write "<tr><td valign='middle' align='center'>" & ii & _
                             "</td><td valign='middle' align='center'>" & rs("ClientName") & "</td><td valign='middle' align='center'>" & rs("ContractID") & "</td><td valign='middle' align='left'>" & rs("GoodsName") & _
                             "</td><td valign='middle' align='center'>" & rs("Specs") & "</td>" & _
                             "<td valign='middle' align='right'>" & gs(zjs1,2) & "/" & rs("AllNumber") & "</td><td valign='middle' align='right'>" & gs(zw1,4) & "</td><td valign='middle' align='center'>" & mid(rs("Date1"),3) & "</td><td valign='middle' align='center'>" & show1(rs("remark")) & "</td></tr>"
        else
		      if zjs1>0 and zw1>0 then
			      response.write "<tr><td valign='middle' align='center'>" & ii & _
                             "</td><td valign='middle' align='center'>" & rs("ClientName") & "</td><td valign='middle' align='center'>" & rs("ContractID") & "</td><td valign='middle' align='left'>" & rs("GoodsName") & _
                             "</td><td valign='middle' align='center'>" & rs("Specs") & "</td>" & _
                             "<td valign='middle' align='right'>" & gs(zjs1,2) & "/" & rs("AllNumber") & "</td><td valign='middle' align='right'>" & gs(zw1,4) & "</td><td valign='middle' align='center'>" & mid(rs("Date1"),3) & "</td><td valign='middle' align='center'>" & show1(rs("remark")) & "</td></tr>"			      
			  end if
		end if					
		
        ktv=ktv+1
		
		if ktv=colsed then
		   response.write "</table></p><p>"
		   
		   response.write "<table width='1000'><tr><td valign='middle' align='center'><b><font size='3'>�ܽ�����</font></b></td></tr></table>" 
    
           response.write "<table border='1' width='1000' bordercolor='#000000'style='border-collapse: collapse' cellspacing='0' cellpadding='0'>"
    
           response.write "<tr><td valign='middle' align='center'>���</td><td valign='middle' align='center'>��������</td><td valign='middle' align='center'><a href='cremainty1.asp?jy=�鿴���&na=ContractID&ye=" & trim(request("ye")) & "&mon=" & trim(request("mon")) & "&ri=" & trim(request("ri")) & "'>��ͬ��</a></td><td valign='middle' align='center'><a href='cremainty1.asp?jy=�鿴���&na=GoodsName&ye=" & trim(request("ye")) & "&mon=" & trim(request("mon")) & "&ri=" & trim(request("ri")) & "'>��Ʒ��</a></td><td valign='middle' align='center'><a href='cremainty1.asp?jy=�鿴���&na=Specs&ye=" & trim(request("ye")) & "&mon=" & trim(request("mon")) & "&ri=" & trim(request("ri")) & "'>���</a></td>" & _
                    "<td valign='middle' align='center'><a href='cremainty1.asp?jy=�鿴���&na=remainnumber&ye=" & trim(request("ye")) & "&mon=" & trim(request("mon")) & "&ri=" & trim(request("ri")) & "'>Ŀǰ�������/����</a></td><td valign='middle' align='center'><a href='cremainty1.asp?jy=�鿴���&na=remainweight&ye=" & trim(request("ye")) & "&mon=" & trim(request("mon")) & "&ri=" & trim(request("ri")) & "'>Ŀǰ������</a></td><td valign='middle' align='center'><a href='cremainty1.asp?jy=�鿴���&na=date1&ye=" & trim(request("ye")) & "&mon=" & trim(request("mon")) & "&ri=" & trim(request("ri")) & "'>����ʱ��</a></td><td valign='middle' align='center'>��ע</td></tr>"

		   ktv=0
		end if
        
        zremainw=zremainw + zw1
        zremainjian=zremainjian + zjs1
        
        rs.movenext  
     loop
 
      rs.close
      set rs=nothing
     
     ' if(cdbl(zremainw)<1) and (cdbl(zremainw)>0) then 
     '       zremainw1= "0" & zremainw1 
     ' end if
            
     response.write "<tr><td valign='middle' align='center' colspan='4'>&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;</td><td valign='middle' align='right'>" & gs(zremainjian,4) & "</td>" & _
                    "<td valign='middle' align='right'>" & gs(zremainw,4) & "</td><td valign='middle' align='center'  colspan='2'>&nbsp;&nbsp;</td></tr>"
     ktv=ktv+1
	 if ktv=colsed then
		   response.write "</table></p><p>"
		   
		   response.write "<table width='1000'><tr><td valign='middle' align='center'><b><font size='3'>�ܽ�����</font></b></td></tr></table>" 
    
           response.write "<table border='1' width='1000' bordercolor='#000000'style='border-collapse: collapse' cellspacing='0' cellpadding='0'>"
    
           response.write "<tr><td valign='middle' align='center'>���</td><td valign='middle' align='center'>��������</td><td valign='middle' align='center'><a href='cremainty1.asp?jy=�鿴���&na=ContractID&ye=" & trim(request("ye")) & "&mon=" & trim(request("mon")) & "&ri=" & trim(request("ri")) & "'>��ͬ��</a></td><td valign='middle' align='center'><a href='cremainty1.asp?jy=�鿴���&na=GoodsName&ye=" & trim(request("ye")) & "&mon=" & trim(request("mon")) & "&ri=" & trim(request("ri")) & "'>��Ʒ��</a></td><td valign='middle' align='center'><a href='cremainty1.asp?jy=�鿴���&na=Specs&ye=" & trim(request("ye")) & "&mon=" & trim(request("mon")) & "&ri=" & trim(request("ri")) & "'>���</a></td>" & _
                    "<td valign='middle' align='center'><a href='cremainty1.asp?jy=�鿴���&na=remainnumber&ye=" & trim(request("ye")) & "&mon=" & trim(request("mon")) & "&ri=" & trim(request("ri")) & "'>Ŀǰ�������/����</a></td><td valign='middle' align='center'><a href='cremainty1.asp?jy=�鿴���&na=remainweight&ye=" & trim(request("ye")) & "&mon=" & trim(request("mon")) & "&ri=" & trim(request("ri")) & "'>Ŀǰ������</a></td><td valign='middle' align='center'><a href='cremainty1.asp?jy=�鿴���&na=date1&ye=" & trim(request("ye")) & "&mon=" & trim(request("mon")) & "&ri=" & trim(request("ri")) & "'>����ʱ��</a></td><td valign='middle' align='center'>��ע</td></tr>"

		   ktv=0
	 end if               
     'response.write "<tr><td valign='middle' align='center' colspan='3'>&nbsp;&nbsp;���²����ܼ�&nbsp;&nbsp;</td><td valign='middle' align='center' colspan='6'>��" & gs(zprices,2) & "</td></tr>"
     response.write "</table><br>"
       
     response.write "<table width='900'><tr><td valign='middle' align='center'>���_________________&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & "�ֿ�(����)��_________________</td></tr></table>"  
     
 end if
 conn.close
 set conn=nothing
%>