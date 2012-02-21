<!-- #include file="data\config.asp" -->
<HTML>
<HEAD>
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<title>结存情况</title>
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
 if trim(request("gsid"))<>"" and request("jy")="查看结存" and trim(request("ye"))<>"" and trim(request("mon"))<>"" and trim(request("ri"))<>"" then
     
	 if not isdate(trim(request("ye")) & "-" & trim(request("mon")) & "-" & trim(request("ri"))) then
        conn.close
        set conn=nothing
        response.write "日期出错！例如：有的月份是31日，有的月份是30日，有的月份是29日，有的月份是28日！"
        response.end  
     end if
	 
	 if cdate(trim(request("ye")) & "-" & trim(request("mon")) & "-" & trim(request("ri"))) > date() then
         response.write "你[选择的时间]大于[当前时间]！"	    
		 conn.close
		 set conn=nothing
		 response.end
     end if
	 
	 tmpday1=cdate(trim(request("ye")) & "-" & trim(request("mon")) & "-" & trim(request("ri")))
	 
     Set rs =server.createobject("adodb.recordset") 
	 
	 Set rs1 =server.createobject("adodb.recordset") 
     
     sql="select ClientName from ClientInfo where ID=" & trim(request("gsid"))
     
     rs.open sql,conn,3,3
     
     tmpname=rs("ClientName")

     rs.close
     
     if request("na")="" then
        'if request("lin")="zero" then
        '  sql="select ContractID,GoodsName,Specs,remainnumber,remainweight,Date1,Time1,remark from InStore where ClientId=" & trim(request("gsid")) & " and remainnumber>0 and remainweight>0 order by Date1 asc,Time1 asc"
        'else  
          sql="select RecNo,BillNo,ContractID,GoodsName,Specs,remainnumber,remainpack,remainweight,Date1,Time1,remark from InStore where ClientId=" & trim(request("gsid")) & " and date1<=#" & tmpday1 & "# order by Date1 asc,Time1 asc"
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
         sql="select RecNo,BillNo,ContractID,GoodsName,Specs,remainnumber,remainpack,remainweight,Date1,Time1,remark from InStore where ClientId=" & trim(request("gsid")) & " and date1<=#" & tmpday1 & "# order by " & request("na") & " " & session("ord") & ",Date1 asc,Time1 asc"
         'end if
 
     end if


     rs.open sql,conn,3,3
      response.write "<p>"  
     response.write "<table width='1000'><tr><td valign='middle' align='center'><b><font size='3'>【" & tmpname & "】结存情况(" & tmpday1 & ")</font></b></td></tr></table>" 
    
     response.write "<table border='1' width='1000' bordercolor='#000000'style='border-collapse: collapse' cellspacing='0' cellpadding='0'>"

      
     response.write "<tr><td valign='middle' align='center'><a href='cremainty.asp?gsid=" & request("gsid") & "&jy=查看结存&na=BillNo&ye=" & trim(request("ye")) & "&mon=" & trim(request("mon")) & "&ri=" & trim(request("ri")) & "'>进仓单号</a></td><td valign='middle' align='center'><a href='cremainty.asp?gsid=" & request("gsid") & "&jy=查看结存&na=ContractID&ye=" & trim(request("ye")) & "&mon=" & trim(request("mon")) & "&ri=" & trim(request("ri")) & "'>合同号</a></td><td valign='middle' align='center'><a href='cremainty.asp?gsid=" & request("gsid") & "&jy=查看结存&na=GoodsName&ye=" & trim(request("ye")) & "&mon=" & trim(request("mon")) & "&ri=" & trim(request("ri")) & "'>产品名</a></td><td valign='middle' align='center'><a href='cremainty.asp?gsid=" & request("gsid") & "&jy=查看结存&na=Specs&ye=" & trim(request("ye")) & "&mon=" & trim(request("mon")) & "&ri=" & trim(request("ri")) & "'>规格</a></td><td valign='middle' align='center'><font color='#0000FF'>件/小件</font></td>" & _
                    "<td valign='middle' align='center'><a href='cremainty.asp?gsid=" & request("gsid") & "&jy=查看结存&na=remainnumber&ye=" & trim(request("ye")) & "&mon=" & trim(request("mon")) & "&ri=" & trim(request("ri")) & "'>结存大件</a></td><td valign='middle' align='center'><a href='cremainty.asp?gsid=" & request("gsid") & "&jy=查看结存&na=remainpack&ye=" & trim(request("ye")) & "&mon=" & trim(request("mon")) & "&ri=" & trim(request("ri")) & "'>结存小捆</a></td><td valign='middle' align='center'><a href='cremainty.asp?gsid=" & request("gsid") & "&jy=查看结存&na=remainweight&ye=" & trim(request("ye")) & "&mon=" & trim(request("mon")) & "&ri=" & trim(request("ri")) & "'>结存吨数</a></td><td valign='middle' align='center'><a href='cremainty.asp?gsid=" & request("gsid") & "&jy=查看结存&na=date1&ye=" & trim(request("ye")) & "&mon=" & trim(request("mon")) & "&ri=" & trim(request("ri")) & "'>进货时间</a></td><td valign='middle' align='center'>备注</td></tr>"
          
     ii=0
        
        zremainw=0
        zremainjian=0
		zremainxbs=0
        
		 
        zprices=0
         
		ktv=0
		colsed=10000     
     do while not rs.eof        
        ii=ii+1
        
        'if (cdbl(rs("remainweight"))<1) and (cdbl(rs("remainweight"))>0) then 
        '    remainw= "0" & rs("remainweight")
        'else
        'remainw=rs("remainweight")    
        'end if
        sql1="select sum(Number1) as zjs,sum(pack1) as zxbs,sum(weight) as zw from OutStore where ClientId=" & trim(request("gsid")) & " and InStoreRecNo=" & rs("RecNo") & " and Date2>#" & tmpday1 & "#"
		
		rs1.open sql1,conn,3,3
		
		if rs1.recordcount>0 then
		   if not isnull(rs1("zjs")) then
		     zjs1=cdbl(rs1("zjs"))+cdbl(rs("remainnumber"))
		   else
		     zjs1=cdbl(rs("remainnumber")) 	 
		   end if
		   
		   if not isnull(rs1("zxbs")) then
		     zxbs1=rs1("zxbs")+rs("remainpack")
		   else
		     zxbs1=rs("remainpack")	 
		   end if
		   
		   if not isnull(rs1("zw")) then
		     zw1=cdbl(rs1("zw"))+cdbl(rs("remainweight"))
		   else
		     zw1=cdbl(rs("remainweight"))	 
		   end if
		else
		   zjs1=cdbl(rs("remainnumber"))
		   zxbs1=cdbl(rs("remainpack"))
		   zw1=cdbl(rs("remainweight"))	   
		end if
		
		rs1.close
		
	    if request("lin")="zero" then
	    		if zjs1=0 then
	    			tempa=gs(0,2)
	    		else
	    			tempa=gs(zxbs1/zjs1,2)
	    		end if
              response.write "<tr><td valign='middle' align='center'>" & rs("BillNo") & _
                             "</td><td valign='middle' align='left'>" & rs("ContractID") & "</td><td valign='middle' align='left'>" & rs("GoodsName") & _
                             "</td><td valign='middle' align='left'>" & rs("Specs") & "</td><td valign='middle' align='right'>" & tempa & "</td>" & _ 
                             "<td valign='middle' align='right'>" & gs(zjs1,2) & "</td><td valign='middle' align='right'>" & gs(zxbs1,2) & "</td><td valign='middle' align='right'>" & gs(zw1,4) & "</td><td valign='middle' align='center'>" & mid(rs("Date1"),3) & "</td><td valign='middle' align='left'>" & show1(rs("remark")) & "</td></tr>"
        else
		      if zjs1>0 and zw1>0 then
			      response.write "<tr><td valign='middle' align='center'>" & rs("BillNo") & _
                             "</td><td valign='middle' align='left'>" & rs("ContractID") & "</td><td valign='middle' align='left'>" & rs("GoodsName") & _
                             "</td><td valign='middle' align='left'>" & rs("Specs") & "</td><td valign='middle' align='right'>" & gs(zxbs1/zjs1,2) & "</td>" & _
                             "<td valign='middle' align='right'>" & gs(zjs1,2) & "</td><td valign='middle' align='right'>" & gs(zxbs1,2) & "</td><td valign='middle' align='right'>" & gs(zw1,4) & "</td><td valign='middle' align='center'>" & mid(rs("Date1"),3) & "</td><td valign='middle' align='left'>" & show1(rs("remark")) & "</td></tr>"			      
			  end if
		end if					
		
        ktv=ktv+1
		
		if ktv=colsed then
		   response.write "</table></p><p>"
		   
		   response.write "<table width='1000'><tr><td valign='middle' align='center'><b><font size='3'>【" & tmpname & "】结存情况</font></b></td></tr></table>" 
    
           response.write "<table border='1' width='1000' bordercolor='#000000'style='border-collapse: collapse' cellspacing='0' cellpadding='0'>"
    
           response.write "<tr><td valign='middle' align='center'><a href='cremainty.asp?gsid=" & request("gsid") & "&jy=查看结存&na=BillNo&ye=" & trim(request("ye")) & "&mon=" & trim(request("mon")) & "&ri=" & trim(request("ri")) & "'>进仓单号</a></td><td valign='middle' align='center'><a href='cremainty.asp?gsid=" & request("gsid") & "&jy=查看结存&na=ContractID&ye=" & trim(request("ye")) & "&mon=" & trim(request("mon")) & "&ri=" & trim(request("ri")) & "'>合同号</a></td><td valign='middle' align='center'><a href='cremainty.asp?gsid=" & request("gsid") & "&jy=查看结存&na=GoodsName&ye=" & trim(request("ye")) & "&mon=" & trim(request("mon")) & "&ri=" & trim(request("ri")) & "'>产品名</a></td><td valign='middle' align='center'><a href='cremainty.asp?gsid=" & request("gsid") & "&jy=查看结存&na=Specs&ye=" & trim(request("ye")) & "&mon=" & trim(request("mon")) & "&ri=" & trim(request("ri")) & "'>规格</a></td><td valign='middle' align='center'><font color='#0000FF'>件/小件</font></td>" & _
                          "<td valign='middle' align='center'><a href='cremainty.asp?gsid=" & request("gsid") & "&jy=查看结存&na=remainnumber&ye=" & trim(request("ye")) & "&mon=" & trim(request("mon")) & "&ri=" & trim(request("ri")) & "'>结存大件</a></td><td valign='middle' align='center'><a href='cremainty.asp?gsid=" & request("gsid") & "&jy=查看结存&na=remainpack&ye=" & trim(request("ye")) & "&mon=" & trim(request("mon")) & "&ri=" & trim(request("ri")) & "'>结存小捆</a></td><td valign='middle' align='center'><a href='cremainty.asp?gsid=" & request("gsid") & "&jy=查看结存&na=remainweight&ye=" & trim(request("ye")) & "&mon=" & trim(request("mon")) & "&ri=" & trim(request("ri")) & "'>结存吨数</a></td><td valign='middle' align='center'><a href='cremainty.asp?gsid=" & request("gsid") & "&jy=查看结存&na=date1&ye=" & trim(request("ye")) & "&mon=" & trim(request("mon")) & "&ri=" & trim(request("ri")) & "'>进货时间</a></td><td valign='middle' align='center'>备注</td></tr>"

		   ktv=0
		end if
        
        zremainw=zremainw + gs(zw1,4)
        zremainjian=zremainjian + gs(zjs1,2)
		zremainxbs=zremainxbs+gs(zxbs1,2)
        
        rs.movenext  
     loop
 
      rs.close
      set rs=nothing
     
     ' if(cdbl(zremainw)<1) and (cdbl(zremainw)>0) then 
     '       zremainw1= "0" & zremainw1 
     ' end if
            
     response.write "<tr><td valign='middle' align='center' colspan='5'>&nbsp;&nbsp;合&nbsp;&nbsp;计&nbsp;&nbsp;</td><td valign='middle' align='right'>" & gs(zremainjian,2) & "</td><td valign='middle' align='right'>" & zremainxbs & "</td>" & _
                    "<td valign='middle' align='right'>" & gs(zremainw,4) & "</td><td valign='middle' align='center'  colspan='2'>&nbsp;&nbsp;</td></tr>"
     ktv=ktv+1
	 if ktv=colsed then
		   response.write "</table></p><p>"
		   
		   response.write "<table width='1000'><tr><td valign='middle' align='center'><b><font size='3'>【" & tmpname & "】结存情况</font></b></td></tr></table>" 
    
           response.write "<table border='1' width='1000' bordercolor='#000000'style='border-collapse: collapse' cellspacing='0' cellpadding='0'>"
    
           response.write "<tr><td valign='middle' align='center'><a href='cremainty.asp?gsid=" & request("gsid") & "&jy=查看结存&na=BillNo&ye=" & trim(request("ye")) & "&mon=" & trim(request("mon")) & "&ri=" & trim(request("ri")) & "'>进仓单号</a></td><td valign='middle' align='center'><a href='cremainty.asp?gsid=" & request("gsid") & "&jy=查看结存&na=ContractID&ye=" & trim(request("ye")) & "&mon=" & trim(request("mon")) & "&ri=" & trim(request("ri")) & "'>合同号</a></td><td valign='middle' align='center'><a href='cremainty.asp?gsid=" & request("gsid") & "&jy=查看结存&na=GoodsName&ye=" & trim(request("ye")) & "&mon=" & trim(request("mon")) & "&ri=" & trim(request("ri")) & "'>产品名</a></td><td valign='middle' align='center'><a href='cremainty.asp?gsid=" & request("gsid") & "&jy=查看结存&na=Specs&ye=" & trim(request("ye")) & "&mon=" & trim(request("mon")) & "&ri=" & trim(request("ri")) & "'>规格</a></td><td valign='middle' align='center'><font color='#0000FF'>件/小件</font></td>" & _
                    "<td valign='middle' align='center'><a href='cremainty.asp?gsid=" & request("gsid") & "&jy=查看结存&na=remainnumber&ye=" & trim(request("ye")) & "&mon=" & trim(request("mon")) & "&ri=" & trim(request("ri")) & "'>结存大件</a></td><td valign='middle' align='center'><a href='cremainty.asp?gsid=" & request("gsid") & "&jy=查看结存&na=remainpack&ye=" & trim(request("ye")) & "&mon=" & trim(request("mon")) & "&ri=" & trim(request("ri")) & "'>结存小捆</a></td><td valign='middle' align='center'><a href='cremainty.asp?gsid=" & request("gsid") & "&jy=查看结存&na=remainweight&ye=" & trim(request("ye")) & "&mon=" & trim(request("mon")) & "&ri=" & trim(request("ri")) & "'>结存吨数</a></td><td valign='middle' align='center'><a href='cremainty.asp?gsid=" & request("gsid") & "&jy=查看结存&na=date1&ye=" & trim(request("ye")) & "&mon=" & trim(request("mon")) & "&ri=" & trim(request("ri")) & "'>进货时间</a></td><td valign='middle' align='center'>备注</td></tr>"

		   ktv=0
	 end if               
     'response.write "<tr><td valign='middle' align='center' colspan='3'>&nbsp;&nbsp;本月仓租总计&nbsp;&nbsp;</td><td valign='middle' align='center' colspan='6'>￥" & gs(zprices,2) & "</td></tr>"
     response.write "</table><br>"
       
     response.write "<table width='900'><tr><td valign='middle' align='center'>填表：" & session("username") & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & "仓库(盖章)：_________________</td></tr></table>"  
     
 end if
  
 conn.close
 set conn=nothing
%>