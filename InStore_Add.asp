<!-- #include file="data\config.asp" -->
<%


shifougenggaichukuzhongliang=request("C3")
if shifougenggaichukuzhongliang="" then 
   shifougenggaichukuzhongliang="OFF"
end if

ClientID=Request.Form("d1")
ContractID=trim( Request.Form("tContractID") )
if ContractID="" then
	ContractID=" "
end if

GoodsName=Request.Form("tGoodsName")

Specs=Request.Form("tSpecs")
if Specs="" then 
	Specs=" "
end if

Pack=Request.Form("tPack")
if Pack="" then
	Pack=" "
end if

AllNumber=Request.Form("tNumber")
RemainNumber = AllNumber
AllPack = Request.Form("tAllPack")
RemainPack=AllPack
PackWeight=Request.Form("tPackWeight")
Weight=Request.Form("tWeight")
Gross=Request.Form("tGross")

if Gross="" then
	Gross="0"
end if

Net=Request.Form("tNet")
if Net="" then
	Net="0"
end if

Time1=trim(Cstr(Formatdatetime(Time(),vbLongTime)))

Date1=request("dYear") & "-" & request("dMonth") & "-" & request("dDay")

Place1=Request.Form("tPlace1")
if Place1="" then
	Place1=" "
end if
Place2=Request.Form("tPlace2")
if Place2="" then
	Place2=" "
end if
Place3=Request.Form("tPlace3")
if Place3="" then
	Place3=" "
end if
RemainWeight=Weight
Operater1 =Session("UserName")
Remark = Request("tRemark")
if Remark="" then
   Remark=" "
end if
Deliver=Request("tDeliver")
if Deliver="" then
	Deliver=" "
end if
Consignee=Request("tConsignee")
if Consignee="" then
	Consignee=" "
end if
Driver=Request("tDriver")
if Driver="" then
	Driver=" "
end if
NumberPlate=Request("tNumberPlate")
if NumberPlate="" then
	NumberPlate=" "
end if

BillNo=Request("tBillNo")
if BillNo="" then
	BillNo=" "
end if

Set rs =server.createobject("adodb.recordset")

if Request("Operate") = "Add" then   '增加新记录

	'将进仓记录追加到数据库

  	 ' rs.Open "Select * from InStore where ContractID=" & "'" &ContractID & "'",conn,1,1

	 ' if rs.recordcount>0 then 
     '	 rs.close
	 '    set rs=nothing
     '	 conn.close
	 '    set conn=nothing
     '
     '	 response.write "合同号已经存在，请重新输入！[<a href='javascript:history.back()'>返回</a>]"
     '   response.end
	 'end if
	
	'生成入库单号No
	dim BillNo1
	
	BillDate=left(Date1,8) & "01"   '每月１号
	BillDate1=mid(BillDate,3,2) & mid(BillDate,6,2)
	sql="select * from BillNo where NoType='进仓通知单' and Date1=#" & BillDate & "#"
	rs.open sql,conn,1,1
	if rs.recordcount<1 then
		rs.close
		sql="insert into BillNo(No1,NoType,Date1) values(1,'进仓通知单'," & "'" & BillDate & "')"
        conn.Execute(sql)
		BillNo1=1
	else
	 	BillNo1=rs("No1") + 1
	 	rs.close
		if cint(right(billno,3)) >= billno1 and mid(Billno,2,4) = BillDate1 then 
		 	sql="Update BillNo set No1=" + CStr(BillNo1) + " where NoType='进仓通知单' and Date1=#" & BillDate & "#"     
		 	conn.Execute(sql)
		else
			billno1=cint(right(billno,3))
		end if
	end if

	for i=1 to  3-len(trim(BillNo1))    '格式化进仓单号
		BillNo1= "0" + cstr(BillNo1)
	next
	BillNo= "J" & BillDate1 & BillNo1     '生成型如：J0410XXX的单号,第一位是进仓的第一个拼音字母，然后四位分别为年月,后三位为流水号

	  rs.open "select * from InStore where ClientID=0"
  
	  zd=""    '字段

	   '获得表的字段
	  for i=1 to rs.fields.count-2
    	      zd=zd + rs.fields(i).name + ","
	  next 

	  zd=zd + rs.fields(i).name

	  rs.close
	  set rs= nothing
  
	  sql="insert into InStore(" + zd + ") values('" + ClientID + "','" + ContractID + "','" + GoodsName + "','" +Specs + "','" + _ 
                   Pack + "'," + AllNumber + "," + AllPack + "," + PackWeight +"," +Weight + "," + Gross + "," + Net + ",'" + _
                   Date1 + "','" + Time1 + "','"  + Place1 + "','" + Place2 + "','" + Place3 + "'," + _
                   RemainNumber + ","  + RemainPack + "," + RemainWeight + ",'"  + Operater1 + "','" + Remark + "','" + _
                   BillNo + "','" + Deliver + "','" + Consignee + "','" + Driver + "','" + NumberPlate + "')" 
  
	  'response.write sql    
	  conn.Execute(sql)
	  
	  if Request("chkPrintInStoreTable")="ON" then      '用户选择打印入库通知单
	  	%>
	  		<script language="javascript">

        		 window.open('Print_InStore_Table.asp?BillNo=<% = BillNo %>&ClientID=<%=ClientID %>&PrintDate=<% =Date1 %>','打印入库通知单','toobar=yes,menubar=yes,scrollbars=yes,resizable=yes') 

			</script>
		<%

	 end if
	 
       if Request("Is_fast") = "yes" then 
	   
	     URL = "InStore_fast.asp?BillNo="& BillNo&"&ClientID="&ClientID 

       else 
         URL = "InStore.asp"
       end if

	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
else  '修改、删除记录或返回

	Select Case Request("ButtonName")
	   Case "Save"  
			'更新进仓数据
			Set rs =server.createobject("adodb.recordset")
			sql="select allnumber,packnumber,remainnumber,remainpack from instore where RecNo=" & request("tRecNo")
			rs.open sql,conn,1,1
			
			if cdbl(AllNumber) < ( rs("allNumber")-rs("remainNumber") )   then
				response.write "件数小于已出仓的件数,修改不成功,请重新修改!"
				rs.close
				set rs=nothing
				response.end
			end if
									
			RemainNumber= cstr( cdbl(AllNumber) - ( rs("AllNumber")- rs("RemainNumber") ) )
			RemainPack1 = cdbl(AllPack)-  ( rs("PackNumber") - rs("RemainPack") )
			RemainWeight=cstr( RemainPack1*cdbl( Request("tNet") )/1000 )
			remainPack=cstr(RemainPack1)
			rs.close
			set rs=nothing			 
						 
			sql ="Update InStore " & _
		    	 "Set ClientID=" & ClientID & "," & _ 
		    	 "ContractID='" & ContractID & "'," & _
			     "GoodsName='" & GoodsName & "'," & _
				 "Specs='" & Specs & "'," & _
				 "Pack='" & Pack & "'," & _ 
				 "AllNumber="  & AllNumber & "," & _
				 "PackNumber=" & AllPack & "," & _
				 "PackWeight=" & PackWeight & "," & _
				 "Weight=" & Weight & "," & _
				 "Gross=" & Gross & "," & _
				 "Net=" & Net & "," & _
				 "Date1='" & Date1 & "'," & _
				 "Time1='" & Time1 & "'," & _
				 "Place1='" & Place1 & "'," & _
				 "Place2='" & Place2 & "'," & _
				 "Place3='" & Place3 & "'," & _
				 "RemainNumber=" & RemainNumber & "," & _
				 "RemainPack=" & RemainPack & "," & _
				 "RemainWeight=" & RemainWeight & ","  & _
				 "Remark='" & Remark & "'," & _
				 "BillNo='"  & BillNo &"'," & _
				 "Deliver='"  & Deliver & "'," & _
				 "Consignee='" & Consignee & "'," & _
				 "Driver='" & Driver & "'," & _
				 "NumberPlate='" & NumberPlate & "' " & _ 				 
				 "Where RecNo=" & Request("tRecNo")

			'response.write sql
			conn.Execute(sql) 
			 		         
			'更新出仓数据
			weight = net & " * pack1 /1000" 
        if   	shifougenggaichukuzhongliang="OFF"	then   '判断是否更改出仓重量
                                      
		   sql ="Update OutStore " & _              
		    	 "Set ClientID=" & ClientID & "," & _ 
		    	 "ContractID='" & ContractID & "'," & _
			     "GoodsName='" & GoodsName & "'," & _
				 "Specs='" & Specs & "'," & _
				 "Pack='" & Pack & "'," & _ 
				 "PackWeight=" & PackWeight & "," & _
				 "Weight=" & Weight & "," & _
				 "Gross=" & Gross & "," & _
				 "Net=" & Net & "," & _
				 "Place1='" & Place1 & "'," & _
				 "Place2='" & Place2 & "'," & _
				 "Place3='" & Place3 & "' " & _
				 "Where InStoreRecNo=" & Request("tRecNo")

			'response.write sql
			'response.end
		else
           	 sql ="Update OutStore " & _
		    	 "Set ClientID=" & ClientID & "," & _ 
		    	 "ContractID='" & ContractID & "'," & _
			     "GoodsName='" & GoodsName & "'," & _
				 "Specs='" & Specs & "'," & _
				 "Place1='" & Place1 & "'," & _
				 "Place2='" & Place2 & "'," & _
				 "Place3='" & Place3 & "' " & _
				 "Where InStoreRecNo=" & Request("tRecNo")
			
		end if	
			
			conn.Execute(sql) 
			
	  
	  Case "Delete"      

			RecNo=Request("RecNo")

			if RecNo<>"" then

				'删除前将该记录保存到DelInStore表中
		  		sql = "insert into DelInStore select * from InStore where RecNo= " & RecNo
		  		conn.Execute(sql)
		  		
		  		'删除进仓记录
		  		sql="Delete From InStore Where RecNo=" & RecNo
	  			conn.Execute(sql)
	  			

				'删除相关出仓记录前将相关记录保存到DelOutStore表中
		  		sql = "insert into DelOutStore select * from OutStore where InStoreRecNo= " & RecNo
		  		conn.Execute(sql)

	  			'删除出仓相关联记录
	  			sql="delete from outstore where Instorerecno=" & recno
	  			conn.execute(sql)

			end if
				  
	  Case "Return"
	  
   End Select
  
   if Session("varClientID") = "All" then
   		url="InStore_browse_all_frame.asp"
   else
	   	url="InStore_browse_frame.asp"
   end if
	
	
end if

'关闭数据库连接
conn.close
set conn=nothing

%>
<html><head><title>进仓管理</title>
<meta http-equiv=refresh content="0; url=<%=url%>">
</head>

</html>