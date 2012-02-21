<%
if session("userright")<>"仓库主管" and session("userright")<>"经理" then
   response.write "你没有使用的权限！"
   response.end
end if
%>
<!-- #include file="data\config.asp" -->
<HTML>
<HEAD>
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<title>月明细表</title>
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
 if trim(request("gsid"))<>"" and trim(request("gs"))<>"" and trim(request("ye"))<>"" and trim(request("mon"))<>"" then
     
     response.write "<table width='900'><tr><td valign='middle' align='center'><b><font size='3'>" & trim(request("gs")) & "," & trim(request("ye")) & "年" & trim(request("mon")) & "月-->挑拨,仓库,仓租月明细表</font></b></td></tr></table><br>" 
    
     response.write "合同号：<u>" & trim(request("ContractID")) & "</u>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;库存单价(元/(吨.日)：<u>" & trim(request("prices")) & "</u>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;产品名：<u>" & trim(request("GoodsName")) & "</u><br><br>上月/月结存件数：<u>" & request("remainjian1") & "</u>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;上月/月结存吨数：<u>" & request("remianw1") & "</u>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;本月/月结存件数：<u>" & request("remainjian") & "</u>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;本月/月结存吨数：<u>" & request("remianw") & "</u><br><br>"               
     
     response.write "<table border='1' width='900' bordercolor='#000000'style='border-collapse: collapse' cellspacing='0' cellpadding='0'>"
    
     response.write "<tr><td valign='middle' align='center'>提货日期</td><td valign='middle' align='center'>提货件数</td><td valign='middle' align='center'>提货吨数</td>" & _
                    "<td valign='middle' align='center'>结存件数</td><td valign='middle' align='center'>结存吨数</td><td valign='middle' align='center'>起始日期</td>" & _
                    "<td valign='middle' align='center'>天数</td><td valign='middle' align='center'>累计仓租</td><td valign='middle' align='center'>&nbsp;&nbsp;备注&nbsp;&nbsp;</td></tr>"
                   
     
          
     Set rs1 =server.createobject("adodb.recordset")
     Set rs =server.createobject("adodb.recordset")                                   
     
     sql="select remainnumber,remainweight from InStore where ClientId=" & trim(request("gsid"))' & " and ContractID='" & trim(request("ContractID")) & "'"
     rs.open sql,conn,3,3   
     
     remainnumber=rs("remainnumber")
     remainweight=rs("remainweight")
     
     rs.close
     
     sql="select Date1,Time1,Number1,Weight from OutStore where ClientId=" & trim(request("gsid")) & " and InStoreRecNo=" & trim(request("InStoreRecNo")) & " and Date1 like '" & trim(request("ye")) & "-" & trim(request("mon")) & "-%'"
     
     'response.write sql
     rs.open sql,conn,3,3      
   
     zjian=0
     zweight=0
     zpricess=0
           
     do while not rs.eof        
        i=i+1
        sql1="select sum(Number1) as zjs,sum(weight) as zw from OutStore where ClientId=" & trim(request("gsid")) & " and InStoreRecNo=" & trim(request("InStoreRecNo")) & " and (Date1>#" & rs("Date1") & "# or (Date1=#" & rs("Date1") & "# and Time1>#" & rs("Time1") & "#)) group by InStoreRecNo"
        'response.write sql1
        'sql1="select count(Number1) as zjs,count(weight) as zw from OutStore where ClientId=" & trim(request("gsid")) & " and ContractID='" & trim(request("ContractID")) & "'and Date1>#" & rs("Date1") & "# group by ContractID"
        rs1.open sql1,conn,3,3
        if rs1.recordcount>0 then
           remainjian=cdbl(remainnumber)+cdbl(rs1("zjs"))   
           remainw=cdbl(remainweight)+cdbl(rs1("zw"))
        else
           remainjian=remainnumber
           remainw=remainweight
        end if
        rs1.close
         
        sql1="select Date1,Time1 from InStore where ClientId=" & trim(request("gsid")) & " and RecNo=" & trim(request("InStoreRecNo"))' & "'" 
        rs1.open sql1,conn,3,3
       
        tian=DateDiff("d",rs1("Date1"),rs("Date1")) + 1
        
        'response.write tian
        
        pricess=cdbl(trim(request("prices"))) * tian * cdbl(rs("Weight"))
        
        'if cdbl(prices)<1 then 
        '   pricess= "0" & prices
        'else
        '   pricess = prices
        'end if   
        
        response.write "<tr><td valign='middle' align='center'>" & rs("Date1") & " " & rs("Time1") & "</td><td valign='middle' align='right'>" & gs(rs("Number1"),4) & "</td><td valign='middle' align='right'>" & gs(rs("Weight"),4) & "</td>" & _
                       "<td valign='middle' align='right'>" & gs(remainjian,4) & "</td><td valign='middle' align='right'>" & gs(remainw,4) & "</td><td valign='middle' align='center'>" & rs1("Date1") & " " & rs1("Time1") & "</td>" & _
                       "<td valign='middle' align='right'>" & tian & "</td><td valign='middle' align='right'>" & gs(pricess,2) & "</td><td valign='middle' align='center'>&nbsp;&nbsp;</td></tr>"
               
        rs1.close
        
        zjian=zjian+rs("Number1")
        zweight=cdbl(zweight)+cdbl(rs("Weight"))
        zpricess=cdbl(zpricess)+pricess
        
        
        rs.movenext  
     loop
 
     rs.close
     set rs1=nothing
     set rs=nothing
     
     'if cdbl(zpricess)<1 then 
     '      zpricess= "0" & zpricess
     'else
     '      zpricess = zpricess
     'end if 
     
     response.write "<tr><td valign='middle' align='center'>合计(总提货(件数/吨数))</td><td valign='middle' align='right'>" & gs(zjian,4) & "</td>" & _
                    "<td valign='middle' align='right'>" & gs(zweight,4) & "</td><td valign='middle' align='center' colspan='4'>本月仓租总计：</td><td valign='middle' align='center' colspan='2'>" & _
                    "￥" & gs(zpricess,2) & "</td></tr>"      
     
     response.write "</table><br>" 
     
     response.write "<table width='900'><tr><td valign='middle' align='center'>填表：_________________&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & "仓库(盖章)：_________________</td></tr></talbe>"  
    
 end if
 conn.close
 set conn=nothing
%>