<%
if session("userright")<>"仓库主管" and session("userright")<>"经理" then
   response.write "你没有使用的权限！"
   response.end
end if
%>
<!-- #include file="data\config.asp" -->
<HTML>
<HEAD>
<META content="Microsoft FrontPage 5.0" name=GENERATOR>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<title>月汇总表</title>
<STYLE type=text/css>
<!--
A:link {text-decoration:none; color:"#0000ff"}
A:visited {text-decoration:none}
A:hover {text-decoration:underline; color:"#FF0033"}
body { font-family:"Verdana", "Arial", "Helvetica", "sans-serif"; font-size: 9pt}
p {  font-family:"Verdana", "Arial", "Helvetica", "sans-serif"; font-size: 9pt}
P {page-break-after: always}  
td {  font-family: "Verdana", "Arial", "Helvetica", "sans-serif"; font-size: 9pt}
-->
</STYLE>
</HEAD>
<%
	
  Server.ScriptTimeOut=9000

 if trim(request("gsid"))<>"" and trim(request("ye"))<>"" and trim(request("ye1"))<>"" and trim(request("mon"))<>"" and trim(request("mon1"))<>"" and trim(request("ri"))<>"" and trim(request("ri1"))<>"" and request("tj1")="挑拨,仓库,仓租月汇总表" then
     
     if not isdate(trim(request("ye")) & "-" & trim(request("mon")) & "-" & trim(request("ri"))) or not isdate(trim(request("ye1")) & "-" & trim(request("mon1")) & "-" & trim(request("ri1"))) then
        conn.close
        set conn=nothing
        response.write "日期出错！例如：有的月份是31日，有的月份是30日，有的月份是29日，有的月份是28日！"
        response.end  
     end if
     
     if cdate(trim(request("ye")) & "-" & trim(request("mon")) & "-" & trim(request("ri"))) > cdate(trim(request("ye1")) & "-" & trim(request("mon1")) & "-" & trim(request("ri1"))) then
        conn.close
        set conn=nothing
        response.write "[开始日期]不能大于[终止日期]！"
        response.end  
     end if
     
     if cdate(trim(request("ye1")) & "-" & trim(request("mon1")) & "-" & trim(request("ri1"))) > date() then
         response.write "你[选择的时间]大于[当前时间]！"	    
		 conn.close
		 set conn=nothing
		 response.end
     end if
	 
     Set rs1 =server.createobject("adodb.recordset")
      
     Set rs =server.createobject("adodb.recordset") 

     Set rs2 =server.createobject("adodb.recordset") 

     
     rs.open "select ClientName from ClientInfo where ID=" & trim(request("gsid")),conn,3,3
     
     if rs.recordcount=0 then 
        set rs1=nothing
        rs.close
        set rs=nothing
        set conn=nothing
        response.end
     end if
     
     tmpgs=rs("ClientName")
     rs.close
     
     response.write "<p>"
     
     response.write "<table width='1000'><tr><td valign='middle' align='center'><b><font size='3'>仓 租 明 细 表</font></b></td></tr></table>" 
     
     response.write "<table width='1000'><tr><td valign='middle' align='left'><b><font size='2'>单位：【" & tmpgs & "】" & "</font></b></td><td valign='middle' align='right'><b><font size='2'>日期：" & trim(request("ye1")) & "-" & trim(request("mon1")) & "-" & trim(request("ri1")) & "</font></b></td></tr></table>" 
    
     response.write "<table border='1' width='1000' bordercolor='#000000'style='border-collapse: collapse' cellspacing='0' cellpadding='0'>"
     
     response.write "<tr><td valign='middle' align='center'>序号</td><td valign='middle' align='center'>合同号</td><td valign='middle' align='center'>品名/规格</td><td valign='middle' align='center'>提货日期</td><td valign='middle' align='center'>提货件数</td><td valign='middle' align='center'>提货吨数</td>" & _
                    "<td valign='middle' align='center'>收费日期</td><td valign='middle' align='center'>计费(元/天)</td><td valign='middle' align='center'>天数</td><td valign='middle' align='center'>仓租(元)</td><td valign='middle' align='center'>结存件数</td><td valign='middle' align='center'>结存吨数</td>" & _
                    "<td valign='middle' align='center'>进货时间</td><td valign='middle' align='center'>备注</td></tr>"
                                
       
     'mon2=int(trim(request("mon")))'本月
     
     'mon1=(int(trim(request("mon")))+1) mod 12'下个月
     'if mon1=0 then mon1=12
	 
     'if int(trim(request("mon")))=12 then
     '   nian=int(trim(request("ye")))+1
     'else
     '   nian=int(trim(request("ye")))   
     'end if
   
     'tmpday=formatdatetime(trim(request("ye")) & "-" & mon2 & "-1",1)
     
     tmpday=cdate(trim(request("ye")) & "-" & trim(request("mon")) & "-" & trim(request("ri")))
     tmpday1=cdate(trim(request("ye1")) & "-" & trim(request("mon1")) & "-" & trim(request("ri1")))
     
     os=tmpday1-tmpday+1
   
     rs1.open "select price from ClientInfo where id=" & trim(request("gsid")),conn,3,3
        
     prices1=cdbl(rs1("price"))
     prices2="￥" & gs(cdbl(rs1("price")),2)
     rs1.close
 
     sql="select Date1,Time1,RecNo,ContractID,GoodsName,Specs,remainnumber,remainweight,remark from InStore where ClientId=" & trim(request("gsid")) & " and Date1<=#" & tmpday1 & "# order by GoodsName asc,Date1 asc,Time1 asc"

	 rs.open sql,conn,3,3
     
     ii=0
     ktv=0
     'zremainw1=0
     'zremainjian1=0
        
     'zremainw=0
     'zremainjian=0

     'zprices=0
     'response.write rs.recordcount     
     
     coled=38
     
  do while not rs.eof        
     rs2.open "select InStoreRecNo from OutStore where InStoreRecNo=" & rs("RecNo") & " and Date1>=" & tmpday,conn,3,3
	 if rs2.recordcount=0 and cdbl(rs("remainweight"))<=0 then
	    rs2.close
	 else
	    rs2.close
        ii=ii+1
        
        '本月出仓
        rs1.open "select Date1,Time1,Number1,weight from Outstore where ClientId=" & trim(request("gsid")) & " and InStoreRecNo=" & rs("RecNo") & " and Date1>=#" & tmpday & "# and Date1<=#" & tmpday1 & "# order by Date1 asc,Time1 asc",conn,3,3
        
        prices=0
          
        if rs1.recordcount=0 then
           
              sql1="select sum(Number1) as zjs,sum(weight) as zw from OutStore where ClientId=" & trim(request("gsid")) & " and InStoreRecNo=" & rs("RecNo") & " and Date1>#" & tmpday1 & "# group by InStoreRecNo"
              rs2.open sql1,conn,3,3
              
              if rs2.recordcount>0 then'算出结存
              
                 if not isnull(rs2("zjs")) then
                   remainjian=cdbl(rs("remainnumber"))+cdbl(rs2("zjs"))   
                 else
                   remainjian=rs("remainnumber")
                 end if  
                 
                 if not isnull(rs2("zw")) then
                    remainw=cdbl(rs("remainweight"))+cdbl(rs2("zw"))
                 else
                    remainw=rs("remainweight")
                 end if   
                 
              else
			    remainjian=rs("remainnumber")
                 remainw=rs("remainweight")
              end if  
              
              rs2.close
			  
              '打印结存计费
              jcj=gs(remainjian,4)
              jcd= gs(remainw,4)
			  
			  if  cdbl(jcj)<0.000001 then
			  		jcd=gs(0,4)
			  end if
              
              flg=1
              if (rs("date1")>tmpday) then '进货时间大于开始时间，则用出:出货时间-进货时间
                  response.write "<tr><td valign='middle' align='center'>" & ii & "</td><td valign='middle' align='center'>" & rs("ContractID") & "</td><td valign='middle' align='left'>" & rs("GoodsName") & "/" & rs("Specs") & "</td><td valign='middle' align='center'>&nbsp;&nbsp</td><td valign='middle' align='right'>&nbsp;&nbsp</td><td valign='middle' align='right'>&nbsp;&nbsp</td>" & _
                                 "<td valign='middle' align='center'>" & mid(cstr(year(rs("date1"))),3) & "." & month(rs("date1")) & "." & day(rs("date1")) & "-" &  mid(trim(request("ye1")),3) & "." & trim(request("mon1")) & "." & trim(request("ri1")) & "</td><td valign='middle' align='center'>" & prices2 & "</td><td valign='middle' align='center'>" & tmpday1-rs("date1")+1 & "</td><td valign='middle' align='right'>￥" & gs(cdbl(remainw)*(tmpday1-rs("date1")+1)*prices1,2) & "</td><td valign='middle' align='right'>" & jcj & "</td><td valign='middle' align='right'>" & jcd & "</td>" & _
                                 "<td valign='middle' align='center'>" & mid(rs("Date1"),3) & "</td><td valign='middle' align='center'>" & show1(rs("remark")) & "</td></tr>"
                  os1=tmpday1-rs("date1")+1               
              else
                  response.write "<tr><td valign='middle' align='center'>" & ii & "</td><td valign='middle' align='center'>" & rs("ContractID") & "</td><td valign='middle' align='left'>" & rs("GoodsName") & "/" & rs("Specs") & "</td><td valign='middle' align='center'>&nbsp;&nbsp</td><td valign='middle' align='right'>&nbsp;&nbsp</td><td valign='middle' align='right'>&nbsp;&nbsp</td>" & _
                                 "<td valign='middle' align='center'>" & mid(trim(request("ye")),3) & "." & trim(request("mon")) & "." & trim(request("ri")) & "-" &  mid(trim(request("ye1")),3) & "." & trim(request("mon1")) & "." & trim(request("ri1")) & "</td><td valign='middle' align='center'>" & prices2 & "</td><td valign='middle' align='center'>" & os & "</td><td valign='middle' align='right'>￥" & gs(cdbl(remainw)*os*prices1,2) & "</td><td valign='middle' align='right'>" & jcj & "</td><td valign='middle' align='right'>" & jcd & "</td>" & _
                                 "<td valign='middle' align='center'>" & mid(rs("Date1"),3) & "</td><td valign='middle' align='center'>" & show1(rs("remark")) & "</td></tr>"
                  flg=2               
              end if
              
              ktv=ktv+1'控制换页表头
              'response.write "<br>" & ktv & "<br>" 
              if ktv=coled then
                 response.write "</table>"
                 
                 response.write "</p><p>"
                 
                 response.write "<table width='1000'><tr><td valign='middle' align='center'><b><font size='3'>仓 租 明 细 表</font></b></td></tr></table>" 
    
                 response.write "<table width='1000'><tr><td valign='middle' align='left'><b><font size='2'>单位：【" & tmpgs & "】" & "</font></b></td><td valign='middle' align='right'><b><font size='2'>日期：" & trim(request("ye1")) & "-" & trim(request("mon1")) & "-" & trim(request("ri1")) & "</font></b></td></tr></table>" 
    
                 response.write "<table border='1' width='1000' bordercolor='#000000'style='border-collapse: collapse' cellspacing='0' cellpadding='0'>"
                 
                 response.write "<tr><td valign='middle' align='center'>序号</td><td valign='middle' align='center'>合同号</td><td valign='middle' align='center'>品名/规格</td><td valign='middle' align='center'>提货日期</td><td valign='middle' align='center'>提货件数</td><td valign='middle' align='center'>提货吨数</td>" & _
                                "<td valign='middle' align='center'>收费日期</td><td valign='middle' align='center'>计费(元/天)</td><td valign='middle' align='center'>天数</td><td valign='middle' align='center'>仓租(元)</td><td valign='middle' align='center'>结存件数</td><td valign='middle' align='center'>结存吨数</td>" & _
                                "<td valign='middle' align='center'>进货时间</td><td valign='middle' align='center'>备注</td></tr>"
                 
                                   
   
                 ktv=0
              end if
      
        else
           
              if (rs("date1")>tmpday) then '进货时间大于开始时间，则用出:出货时间-进货时间
                   tmppr=(rs1("Date1")-rs("date1")+1)*prices1*cdbl(rs1("weight"))'算出仓的租金 
              else
                   tmppr=(rs1("Date1")-tmpday+1)*prices1*cdbl(rs1("weight"))'算出仓的租金               
              end if

              sql1="select sum(Number1) as zjs,sum(weight) as zw from OutStore where ClientId=" & trim(request("gsid")) & " and InStoreRecNo=" & rs("RecNo") & " and (Date1>#" & rs1("Date1") & "# or (Date1=#" & rs1("Date1") & "# and Time1>#" & rs1("Time1") & "#)) group by InStoreRecNo"
              rs2.open sql1,conn,3,3
            
              if rs2.recordcount>0 then'算出结存
              
                if not isnull(rs2("zjs")) then
                   remainjian=cdbl(rs("remainnumber"))+cdbl(rs2("zjs"))   
                else
                   remainjian=rs("remainnumber")
                end if   
                
                if not isnull(rs2("zw")) then
                   remainw=cdbl(rs("remainweight"))+cdbl(rs2("zw"))
                else
                   remainw=rs("remainweight")
                end if   
                
              else
                 remainjian=rs("remainnumber")
                 remainw=rs("remainweight")
              end if
              rs2.close
               
              '第一次，打印出仓记录 
              jcj=gs(remainjian,4)
              jcd= gs(remainw,4)
              
              if (rs("date1")>tmpday) then '进货时间大于开始时间，则用出:出货时间-进货时间
                  response.write "<tr><td valign='middle' align='center' rowspan='" & rs1.recordcount+1 & "'>" & ii & "</td><td valign='middle' align='center' rowspan='" & rs1.recordcount+1 & "'>" & rs("ContractID") & "</td><td valign='middle' align='left' rowspan='" & rs1.recordcount+1 & "'>" & rs("GoodsName") & "/" & rs("Specs")  & "</td><td valign='middle' align='center'>" & mid(rs1("Date1"),3) & "</td><td valign='middle' align='right'>" & rs1("Number1") & "</td><td valign='middle' align='right'>" & gs(rs1("weight"),4) & "</td>" & _
                             "<td valign='middle' align='center'>" & mid(cstr(year(rs("date1"))),3) & "." & month(rs("date1")) & "." & day(rs("date1")) & "-" & mid(cstr(year(rs1("date1"))),3) & "." & month(rs1("date1")) & "." & day(rs1("date1")) & "</td><td valign='middle' align='center'>" & prices2 & "</td><td valign='middle' align='center'>" & rs1("Date1")-rs("Date1")+1 & "</td><td valign='middle' align='right'>￥" & gs(tmppr,2) & "</td><td valign='middle' align='right'>" & jcj & "</td><td valign='middle' align='right'>" & jcd & "</td>" & _
                             "<td valign='middle' align='center' rowspan='" & rs1.recordcount+1 & "'>" & mid(rs("Date1"),3) & "</td><td valign='middle' align='center' rowspan='" & rs1.recordcount+1 & "'>" & show1(rs("remark")) & "</td></tr>"
              else 
                  response.write "<tr><td valign='middle' align='center' rowspan='" & rs1.recordcount+1 & "'>" & ii & "</td><td valign='middle' align='center' rowspan='" & rs1.recordcount+1 & "'>" & rs("ContractID") & "</td><td valign='middle' align='left' rowspan='" & rs1.recordcount+1 & "'>" & rs("GoodsName") & "/" & rs("Specs")  & "</td><td valign='middle' align='center'>" & mid(rs1("Date1"),3) & "</td><td valign='middle' align='right'>" & rs1("Number1") & "</td><td valign='middle' align='right'>" & gs(rs1("weight"),4) & "</td>" & _
                             "<td valign='middle' align='center'>" & mid(trim(request("ye")),3) & "." & trim(request("mon")) & "." & trim(request("ri")) & "-" & mid(cstr(year(rs1("date1"))),3) & "." & month(rs1("date1")) & "." & day(rs1("date1")) & "</td><td valign='middle' align='center'>" & prices2 & "</td><td valign='middle' align='center'>" & rs1("Date1")-tmpday+1 & "</td><td valign='middle' align='right'>￥" & gs(tmppr,2) & "</td><td valign='middle' align='right'>" & jcj & "</td><td valign='middle' align='right'>" & jcd & "</td>" & _
                             "<td valign='middle' align='center' rowspan='" & rs1.recordcount+1 & "'>" & mid(rs("Date1"),3) & "</td><td valign='middle' align='center' rowspan='" & rs1.recordcount+1 & "'>" & show1(rs("remark")) & "</td></tr>"
              end if
              
              ktv=ktv+1'控制换页表头
              
              if ktv=coled then
                 response.write "</table>"
                 
                 response.write "</p><p>"
                 
                 response.write "<table width='1000'><tr><td valign='middle' align='center'><b><font size='3'>仓 租 明 细 表</font></b></td></tr></table>" 
    
                 response.write "<table width='1000'><tr><td valign='middle' align='left'><b><font size='2'>单位：【" & tmpgs & "】" & "</font></b></td><td valign='middle' align='right'><b><font size='2'>日期：" & trim(request("ye1")) & "-" & trim(request("mon1")) & "-" & trim(request("ri1")) & "</font></b></td></tr></table>" 
    
                 response.write "<table border='1' width='1000' bordercolor='#000000'style='border-collapse: collapse' cellspacing='0' cellpadding='0'>"
                 
                 response.write "<tr><td valign='middle' align='center'>序号</td><td valign='middle' align='center'>合同号</td><td valign='middle' align='center'>品名/规格</td><td valign='middle' align='center'>提货日期</td><td valign='middle' align='center'>提货件数</td><td valign='middle' align='center'>提货吨数</td>" & _
                                "<td valign='middle' align='center'>收费日期</td><td valign='middle' align='center'>计费(元/天)</td><td valign='middle' align='center'>天数</td><td valign='middle' align='center'>仓租(元)</td><td valign='middle' align='center'>结存件数</td><td valign='middle' align='center'>结存吨数</td>" & _
                                "<td valign='middle' align='center'>进货时间</td><td valign='middle' align='center'>备注</td></tr>"
                 
                 ktv=0
              end if
              
              prices=prices+tmppr'累计租金
                    
              rs1.movenext
              
              reco=0
              
              while not rs1.eof
              
                 if (rs("date1")>tmpday) then '进货时间大于开始时间，则用出:出货时间-进货时间
                    tmppr=(rs1("Date1")-rs("date1")+1)*prices1*cdbl(rs1("weight"))'算出仓的租金 
                 else
                    tmppr=(rs1("Date1")-tmpday+1)*prices1*cdbl(rs1("weight"))'算出仓的租金               
                 end if             
              
                 sql1="select sum(Number1) as zjs,sum(weight) as zw from OutStore where ClientId=" & trim(request("gsid")) & " and InStoreRecNo=" & rs("RecNo") & " and (Date1>#" & rs1("Date1") & "# or (Date1=#" & rs1("Date1") & "# and Time1>#" & rs1("Time1") & "#)) group by InStoreRecNo"
                 
                 rs2.open sql1,conn,3,3
                 if rs2.recordcount>0 then'算出结存
              
                    if not isnull(rs2("zjs")) then
                       remainjian=cdbl(rs("remainnumber"))+cdbl(rs2("zjs"))   
                    else
                       remainjian=rs("remainnumber")
                    end if   
                
                    if not isnull(rs2("zw")) then
                       remainw=cdbl(rs("remainweight"))+cdbl(rs2("zw"))
                    else
                       remainw=rs("remainweight")
                    end if   
                
                 else
                    remainjian=rs("remainnumber")
                    remainw=rs("remainweight")
                 end if
                 
                 rs2.close
                 '再次，打印出仓记录
                 jcj=gs(remainjian,4)
                 jcd= gs(remainw,4) 
                 
                 
                 if ktv=0 then   
                     if (rs("date1")>tmpday) then '进货时间大于开始时间，则用出:出货时间-进货时间
                          response.write "<tr><td valign='middle' align='center' rowspan='" & rs1.recordcount-reco & "'>" & ii & "</td><td valign='middle' align='center' rowspan='" & rs1.recordcount-reco & "'>" & rs("ContractID") & "</td><td valign='middle' align='left' rowspan='" & rs1.recordcount-reco & "'>" & rs("GoodsName") & "/" & rs("Specs")  & "</td><td valign='middle' align='center'>" & mid(rs1("Date1"),3) & "</td><td valign='middle' align='right'>" & rs1("Number1") & "</td><td valign='middle' align='right'>" & gs(rs1("weight"),4) & "</td>" & _
                                         "<td valign='middle' align='center'>" & mid(cstr(year(rs("date1"))),3) & "." & month(rs("date1")) & "." & day(rs("date1")) & "-" & mid(cstr(year(rs1("date1"))),3) & "." & month(rs1("date1")) & "." & day(rs1("date1")) & "</td><td valign='middle' align='center'>" & prices2 & "</td><td valign='middle' align='center'>" & rs1("Date1")-rs("Date1")+1 & "</td><td valign='middle' align='right'>￥" & gs(tmppr,2) & "</td><td valign='middle' align='right'>" & jcj & "</td><td valign='middle' align='right'>" & jcd & "</td>" & _
                                         "<td valign='middle' align='center' rowspan='" & rs1.recordcount-reco & "'>" & mid(rs("Date1"),3) & "</td><td valign='middle' align='center' rowspan='" & rs1.recordcount-reco & "'>&nbsp;</td></tr>"
                     else 
                          response.write "<tr><td valign='middle' align='center' rowspan='" & rs1.recordcount-reco & "'>" & ii & "</td><td valign='middle' align='center' rowspan='" & rs1.recordcount-reco & "'>" & rs("ContractID") & "</td><td valign='middle' align='left' rowspan='" & rs1.recordcount-reco & "'>" & rs("GoodsName") & "/" & rs("Specs")  & "</td><td valign='middle' align='center'>" & mid(rs1("Date1"),3) & "</td><td valign='middle' align='right'>" & rs1("Number1") & "</td><td valign='middle' align='right'>" & gs(rs1("weight"),4) & "</td>" & _
                                         "<td valign='middle' align='center'>" & mid(trim(request("ye")),3) & "." & trim(request("mon")) & "." & trim(request("ri")) & "-" & mid(cstr(year(rs1("date1"))),3) & "." & month(rs1("date1")) & "." & day(rs1("date1")) & "</td><td valign='middle' align='center'>" & prices2 & "</td><td valign='middle' align='center'>" & rs1("Date1")-tmpday+1 & "</td><td valign='middle' align='right'>￥" & gs(tmppr,2) & "</td><td valign='middle' align='right'>" & jcj & "</td><td valign='middle' align='right'>" & jcd & "</td>" & _
                                         "<td valign='middle' align='center' rowspan='" & rs1.recordcount-reco & "'>" & mid(rs("Date1"),3) & "</td><td valign='middle' align='center' rowspan='" & rs1.recordcount-reco & "'>&nbsp;</td></tr>"
                     end if
                 
                 else
                 
                     if (rs("date1")>tmpday) then '进货时间大于开始时间，则用出:出货时间-进货时间 
                          response.write "<tr><td valign='middle' align='center'>" & mid(rs1("Date1"),3) & "</td><td valign='middle' align='right'>" & rs1("Number1") & "</td><td valign='middle' align='right'>" & gs(rs1("weight"),4) & "</td>" & _
                                         "<td valign='middle' align='center'>" & mid(cstr(year(rs("date1"))),3) & "." & month(rs("date1")) & "." & day(rs("date1")) & "-" & mid(cstr(year(rs1("date1"))),3) & "." & month(rs1("date1")) & "." & day(rs1("date1")) & "</td><td valign='middle' align='center'>" & prices2 & "</td><td valign='middle' align='center'>" & rs1("Date1")-rs("Date1")+1 & "</td><td valign='middle' align='right'>￥" & gs(tmppr,2) & "</td><td valign='middle' align='right'>" & jcj & "</td><td valign='middle' align='right'>" & jcd & "</td></tr>"
                     else
                          response.write "<tr><td valign='middle' align='center'>" & mid(rs1("Date1"),3) & "</td><td valign='middle' align='right'>" & rs1("Number1") & "</td><td valign='middle' align='right'>" & gs(rs1("weight"),4) & "</td>" & _
                                         "<td valign='middle' align='center'>" & mid(trim(request("ye")),3) & "." & trim(request("mon")) & "." & trim(request("ri")) & "-" & mid(cstr(year(rs1("date1"))),3) & "." & month(rs1("date1")) & "." & day(rs1("date1")) & "</td><td valign='middle' align='center'>" & prices2 & "</td><td valign='middle' align='center'>" & rs1("Date1")-tmpday+1 & "</td><td valign='middle' align='right'>￥" & gs(tmppr,2) & "</td><td valign='middle' align='right'>" & jcj & "</td><td valign='middle' align='right'>" & jcd & "</td></tr>"
                     end if                    
                    
                 end if
                 
   
                 ktv=ktv+1'控制换页表头
                                  
                 if ktv=coled then
                         response.write "</table>"
                 
                         response.write "</p><p>"
                 
                         response.write "<table width='1000'><tr><td valign='middle' align='center'><b><font size='3'>仓 租 明 细 表</font></b></td></tr></table>" 
    
                         response.write "<table width='1000'><tr><td valign='middle' align='left'><b><font size='2'>单位：【" & tmpgs & "】" & "</font></b></td><td valign='middle' align='right'><b><font size='2'>日期：" & trim(request("ye1")) & "-" & trim(request("mon1")) & "-" & trim(request("ri1")) & "</font></b></td></tr></table>" 
    
                         response.write "<table border='1' width='1000' bordercolor='#000000'style='border-collapse: collapse' cellspacing='0' cellpadding='0'>"
                 
                         response.write "<tr><td valign='middle' align='center'>序号</td><td valign='middle' align='center'>合同号</td><td valign='middle' align='center'>品名/规格</td><td valign='middle' align='center'>提货日期</td><td valign='middle' align='center'>提货件数</td><td valign='middle' align='center'>提货吨数</td>" & _
                                  "<td valign='middle' align='center'>收费日期</td><td valign='middle' align='center'>计费(元/天)</td><td valign='middle' align='center'>天数</td><td valign='middle' align='center'>仓租(元)</td><td valign='middle' align='center'>结存件数</td><td valign='middle' align='center'>结存吨数</td>" & _
                                  "<td valign='middle' align='center'>进货时间</td><td valign='middle' align='center'>备注</td></tr>"
                 
                         ktv=0
                 end if 
              
                 prices=prices+tmppr'累计租金
                 
                 rs1.movenext
                 reco=reco+1
              wend
              
              '最后，打印结存计费     
              flg=1
              
              if ktv=0 then
                 if (rs("date1")>tmpday) then '进货时间大于开始时间，则用出:出货时间-进货时间 
                       response.write "<tr><td valign='middle' align='right' colspan='3'>&nbsp;&nbsp;</td><td valign='middle' align='center'>&nbsp;&nbsp;</td><td valign='middle' align='right'>&nbsp;&nbsp;</td><td valign='middle' align='right'>&nbsp;&nbsp;</td>" & _
                                      "<td valign='middle' align='center'>" & mid(cstr(year(rs("date1"))),3) & "." & month(rs("date1")) & "." & day(rs("date1")) & "-" & mid(trim(request("ye1")),3) & "." & trim(request("mon1")) & "." & trim(request("ri1")) & "</td><td valign='middle' align='center'>" & prices2 & "</td><td valign='middle' align='center'>" & tmpday1-rs("date1")+1 & "</td><td valign='middle' align='right'>￥" & gs(cdbl(remainw)*(tmpday1-rs("date1")+1)*prices1,2) & "</td><td valign='middle' align='right'>" & jcj & "</td><td valign='middle' align='right'>" & jcd & "</td><td valign='middle' align='right' colspan='2'>&nbsp;&nbsp;</td></tr>"
                      os1=tmpday1-rs("date1")+1
                 else
                       response.write "<tr><td valign='middle' align='right' colspan='3'>&nbsp;&nbsp;</td><td valign='middle' align='right'>&nbsp;&nbsp;</td><td valign='middle' align='center'>&nbsp;&nbsp;</td><td valign='middle' align='right'>&nbsp;&nbsp;</td>" & _
                                    "<td valign='middle' align='center'>" & mid(trim(request("ye")),3) & "." & trim(request("mon")) & "." & trim(request("ri")) & "-" & mid(trim(request("ye1")),3) & "." & trim(request("mon1")) & "." & trim(request("ri1")) & "</td><td valign='middle' align='center'>" & prices2 & "</td><td valign='middle' align='center'>" & os & "</td><td valign='middle' align='right'>￥" & gs(cdbl(remainw)*os*prices1,2) & "</td><td valign='middle' align='right'>" & jcj & "</td><td valign='middle' align='right'>" & jcd & "</td><td valign='middle' align='right' colspan='2'>&nbsp;&nbsp;</td></tr>"
                      flg=2               
                 end if              
              else
                         
                 if (rs("date1")>tmpday) then '进货时间大于开始时间，则用出:出货时间-进货时间 
                       response.write "<tr><td valign='middle' align='center'>&nbsp;&nbsp;</td><td valign='middle' align='right'>&nbsp;&nbsp;</td><td valign='middle' align='right'>&nbsp;&nbsp;</td>" & _
                                      "<td valign='middle' align='center'>" & mid(cstr(year(rs("date1"))),3) & "." & month(rs("date1")) & "." & day(rs("date1")) & "-" & mid(trim(request("ye1")),3) & "." & trim(request("mon1")) & "." & trim(request("ri1")) & "</td><td valign='middle' align='center'>" & prices2 & "</td><td valign='middle' align='center'>" & tmpday1-rs("date1")+1 & "</td><td valign='middle' align='right'>￥" & gs(cdbl(remainw)*(tmpday1-rs("date1")+1)*prices1,2) & "</td><td valign='middle' align='right'>" & jcj & "</td><td valign='middle' align='right'>" & jcd & "</td></tr>"
                      os1=tmpday1-rs("date1")+1
                 else
                      response.write "<tr><td valign='middle' align='center'>&nbsp;&nbsp;</td><td valign='middle' align='right'>&nbsp;&nbsp;</td><td valign='middle' align='right'>&nbsp;&nbsp;</td>" & _
                                    "<td valign='middle' align='center'>" & mid(trim(request("ye")),3) & "." & trim(request("mon")) & "." & trim(request("ri")) & "-" & mid(trim(request("ye1")),3) & "." & trim(request("mon1")) & "." & trim(request("ri1")) & "</td><td valign='middle' align='center'>" & prices2 & "</td><td valign='middle' align='center'>" & os & "</td><td valign='middle' align='right'>￥" & gs(cdbl(remainw)*os*prices1,2) & "</td><td valign='middle' align='right'>" & jcj & "</td><td valign='middle' align='right'>" & jcd & "</td></tr>"
                      flg=2               
                 end if
                 
              end if
              
              ktv=ktv+1'控制换页表头
              
              if ktv=coled then
                         response.write "</table>"
                 
                         response.write "</p><p>"
                 
                         response.write "<table width='1000'><tr><td valign='middle' align='center'><b><font size='3'>仓 租 明 细 表</font></b></td></tr></table>" 
    
                         response.write "<table width='1000'><tr><td valign='middle' align='left'><b><font size='2'>单位：【" & tmpgs & "】" & "</font></b></td><td valign='middle' align='right'><b><font size='2'>日期：" & trim(request("ye1")) & "-" & trim(request("mon1")) & "-" & trim(request("ri1")) & "</font></b></td></tr></table>" 
    
                         response.write "<table border='1' width='1000' bordercolor='#000000'style='border-collapse: collapse' cellspacing='0' cellpadding='0'>"
                 
                         response.write "<tr><td valign='middle' align='center'>序号</td><td valign='middle' align='center'>合同号</td><td valign='middle' align='center'>品名/规格</td><td valign='middle' align='center'>提货日期</td><td valign='middle' align='center'>提货件数</td><td valign='middle' align='center'>提货吨数</td>" & _
                                  "<td valign='middle' align='center'>收费日期</td><td valign='middle' align='center'>计费(元/天)</td><td valign='middle' align='center'>天数</td><td valign='middle' align='center'>仓租(元)</td><td valign='middle' align='center'>结存件数</td><td valign='middle' align='center'>结存吨数</td>" & _
                                  "<td valign='middle' align='center'>进货时间</td><td valign='middle' align='center'>备注</td></tr>"
                 
                         ktv=0
              end if
              
              
        end if  
           
		if flg=1 then          
               response.write "<tr><td valign='middle' align='center' colspan='9'>合计仓租(元)：</td><td valign='middle' align='right'>￥" & gs(cdbl(remainw)*os1*prices1+prices,2) & "</td><td colspan='5'>&nbsp;</td></tr>"
               zprices=zprices+(cdbl(remainw)*prices1*os1+prices) 
        else
               response.write "<tr><td valign='middle' align='center' colspan='9'>合计仓租(元)：</td><td valign='middle' align='right'>￥" & gs(cdbl(remainw)*os*prices1+prices,2) & "</td><td colspan='5'>&nbsp;</td></tr>"
               zprices=zprices+(cdbl(remainw)*prices1*os+prices) 
        end if
              
           
        rs1.close
        
        
        ktv=ktv+1'控制换页表头
        if ktv=coled then
              response.write "</table>"
                 
              response.write "</p><p>"
                 
              response.write "<table width='1000'><tr><td valign='middle' align='center'><b><font size='3'>仓 租 明 细 表</font></b></td></tr></table>" 
    
              response.write "<table width='1000'><tr><td valign='middle' align='left'><b><font size='2'>单位：【" & tmpgs & "】" & "</font></b></td><td valign='middle' align='right'><b><font size='2'>日期：" & trim(request("ye1")) & "-" & trim(request("mon1")) & "-" & trim(request("ri1")) & "</font></b></td></tr></table>" 
    
              response.write "<table border='1' width='1000' bordercolor='#000000'style='border-collapse: collapse' cellspacing='0' cellpadding='0'>"
                 
              response.write "<tr><td valign='middle' align='center'>序号</td><td valign='middle' align='center'>合同号</td><td valign='middle' align='center'>品名/规格</td><td valign='middle' align='center'>提货日期</td><td valign='middle' align='center'>提货件数</td><td valign='middle' align='center'>提货吨数</td>" & _
                                  "<td valign='middle' align='center'>收费日期</td><td valign='middle' align='center'>计费(元/天)</td><td valign='middle' align='center'>天数</td><td valign='middle' align='center'>仓租(元)</td><td valign='middle' align='center'>结存件数</td><td valign='middle' align='center'>结存吨数</td>" & _
                                  "<td valign='middle' align='center'>进货时间</td><td valign='middle' align='center'>备注</td></tr>"
                 
              ktv=0
         end if
		
     end if
	 
	 
	  
     rs.movenext  
       
  loop
 
     rs.close
     set rs1=nothing
     set rs2=nothing
     set rs=nothing
                          
     response.write "<tr><td valign='middle' align='center' colspan='9'>&nbsp;&nbsp;本期仓租总计&nbsp;&nbsp;</td><td valign='middle' align='center' colspan='1'>￥" & gs(zprices,2) & "</td><td valign='middle' align='center' colspan='5'>&nbsp;</td></tr>"
     response.write "</table><br>"
       
     response.write "<table width='1000'><tr><td valign='middle' align='center'>填表：" & session("username") & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & "仓库(盖章)：_________________</td></tr></table></p>"  
     
 end if
 
 conn.close
 set conn=nothing
%>