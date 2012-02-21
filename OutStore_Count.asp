<!-- #include file="data\config.asp" -->

<%
RecNo=CStr(Request("tRecNo"))

ClientID=Request("d1")
ContractID=trim( Request("tContractID") )
if ContractID="" then
	ContractID=" "
end if

GoodsName=Request("tGoodsName")
Specs=Request("tSpecs")
if Specs="" then
	Specs=" "
end if

Pack=Request("tPack")
if Pack="" then
	Pack=" "
end if

AllNumber=Request("tOutNumber")
AllPack = Request("tOutPack")
PackWeight=Request("tPackWeight")
Weight=Request("tOutWeight")
Gross=Request("tGross")
if Gross="" then
	Gross="0"
end if

if request("chkWeight")="ON" then
	lockweight=TRUE
else
	lockweight=false
end if

Net=Request("tNet")
if Net="" then
	Net="0"
end if

''''''''''''''''''''''''''''''add

    XIANSHI=""
    if request("c2")="ON" then
        XIANSHI="ON"
	ELSE
	    XIANSHI="OFF"
	END IF
''''''''''''''''''''''''''''''end

Time1=trim(Cstr(Formatdatetime(Time(),vbLongTime)))

'格式化计费日期
Date1=trim(Request("tYear")) & "-"

if len(trim(Request("tMonth")))=1 then 
   Date1=Date1 & "0" & trim(Request("tMonth")) & "-"
else
   Date1=Date1 & trim(Request("tmonth")) & "-"
end if 

if len(trim(Request("tday")))=1 then 
   Date1=Date1 & "0" & trim(Request("tDay"))
else
   Date1=Date1 & trim(Request("tDay"))
end if 


'格式化操作日期
Date2=trim(Request("tSYear")) & "-"

if len(trim(Request("tSMonth")))=1 then 
   Date2=Date2 & "0" & trim(Request("tSMonth")) & "-"
else
   Date2=Date2 & trim(Request("tsmonth")) & "-"
end if 

if len(trim(Request("tsday")))=1 then 
   Date2=Date2 & "0" & trim(Request("tSDay"))
else
   Date2=Date2 & trim(Request("tsDay"))
end if 


Place1=Request("tPlace1")
if Place1="" then
	Place1=" "
end if
Place2=Request("tPlace2")
if Place2="" then
	Place2=" "
end if
Place3=Request("tPlace3")
if Place3="" then
	Place3=" "
end if

PickUpID=Request("tPickUpID")
if PickUpID="" then
	PickUpID=" "
end if

Resume1=" "

Operater1 =Session("UserName")
Remark = Request("tOutRemark")
if Remark="" then
   Remark=" "
end if

BillNo1="0"

StartAddress=request("tStartAddress")
if StartAddress="" then
	StartAddress=" "
end if

NumberPlate=request("tNumberPlate")
if NumberPlate="" then
	NumberPlate=" "
end if

Crop=request("tCrop")
if Crop="" then
	Crop=" "
end if

Consigner=request("tConsigner")
if Consigner="" then
	Consigner=" "
end if


Set rs =server.createobject("adodb.recordset")

if Request("Operate") = "Add" then   '增加新记录

	'生成出库单号No
	BillNo=cint(right(trim(request("tBillNo")),3))
	
	
	BillDate=left(Date2,8) & "01"   '每月１号
	
	sql="select * from BillNo where NoType='出仓通知单' and Date1=#" & BillDate & "#"
	rs.open sql,conn,1,1
	if rs.recordcount<1 then
		rs.close
		sql="insert into BillNo(No1,NoType,Date1) values(1,'出仓通知单'," & "'" & BillDate & "')"
        conn.Execute(sql)
		BillNo=1
	else
	 	if BillNo = rs("No1") + 1 then
	 		rs.close
	 		sql="Update BillNo set No1=" + CStr(BillNo) + " where NoType='出仓通知单' and Date1=#" & BillDate & "#"     
		 	conn.Execute(sql)
		else
			rs.close
		end if
		
	end if

	for i=1 to  3-len(trim(BillNo))    '格式化出仓单号
		BillNo= "0" + cstr(BillNo)
	next
	BillNo="C" & mid(BillDate,3,2) & mid(BillDate,6,2) & BillNo     '生成型如：C0410XXX的单号,第一位为出仓拼音，然后四位分别为年月,后三位为流水号

	
	
	
	
	'将进仓记录追加到数据库
      sql="Select * from outStore where PickUpID=" & "'" & PickUpID & "'"
  	  rs.Open sql,conn,1,1
	 ' if rs.recordcount > 0 then 
     '	 rs.close
	 '    set rs=nothing
     ' 	 conn.close
	 '     set conn=nothing
     '
     '	 response.write "提单号已经存在，请重新输入！[<a href='javascript:history.back()'>返回</a>]"
     '    response.end
	 ' end if
	  
	  zd=""    '字段

	   '获得表的字段
	  for i=1 to rs.fields.count-2
    	      zd=zd + rs.fields(i).name + ","
	  next 

	  zd=zd + rs.fields(i).name

	  rs.close
	  set rs= nothing
  
	  sql="insert into OutStore(" + zd + ") values(" + RecNo + ",'" + ClientID + "','" + ContractID + "','" + GoodsName + "','" +Specs + "','" + _ 
                   Pack + "'," + AllNumber + "," + AllPack + "," + PackWeight +"," +Weight + "," + Gross + "," + Net + ",'" + _
                   Date1 + "','" + Date2 + "','" + Time1 + "','"  + Place1 + "','" + Place2 + "','" + Place3 + "','" + _
                   Resume1 + "','"  + PickUpID + "','" + Operater1 + "','" + Remark + "','" + BillNo + "','" + BillNo1 + "','" + StartAddress + "','" + _
                   Crop + "','" + NumberPlate + "','" + Consigner + "'," & lockweight & ")" 
  
	  'response.write sql    
	  'response.end
	  conn.Execute(sql)
	  
	  
	  sql="Update InStore " & _
	  	  "Set RemainNumber=" & request("tNumber") - AllNumber & "," & _
	  	  "RemainPack=" & request("tAllPack") - AllPack & "," & _
	  	  "RemainWeight=" &request("tWeight") - Weight & " " & _
	  	  "where RecNo=" & request("tRecNo")
    ' sql="update Outstore set remainNumber=19 where ContractID='dfsd' "
	   'response.write sql
	   conn.Execute(sql)
	  
	  if request("c1")="ON" then     '用户选择打印出仓通知单
%>
	  		<script language="javascript">

        		 window.open('Print_OutStore_Table.asp?ClientID=<%=ClientID %>&BillNo=<%=BillNo%>&PrintDate=<% =Date2 %>&XIANSHI=<%=XIANSHI%>','打印出库通知单','toobar=yes,menubar=yes,scrollbars=yes,resizable=yes') 

			</script>
<%
	  end if
	  
	 if session("varClientID") = "All" then
	 	 URL = "outStore_Browse_All_Frame.asp"
	 else
	 	 URL = "outStore_Browse_Frame.asp"
	 end if
	 
else  '修改、删除记录或返回

	Select Case Request("ButtonName")
	   Case "Save"

       	  ' rs.Open "Select * from InStore where ContractID=" & "'" &ContractID & "'",conn,1,1
		  ' 
		  ' if rs.recordcount>0 then  
    	  '		 rs.close
	      '		 set rs=nothing
	      '		 conn.close
		  '	     set conn=nothing
          '
          '		 response.write "合同号已已经存在，请重新输入！[<a href='javascript:history.back()'>返回</a>]"
	      '    	 response.end
		  '	end if
		  '	rs.close
		  '	set rs=nothing

			sql ="Update InStore " & _
		    	 "Set ContractID=" & "'" & ContractID & "'," & _
			     "GoodsName=" & "'" & GoodsName & "'," & _
				 "Specs=" & "'" & Specs & "'," & _
				 "Pack=" & "'" & Pack & "'," & _ 
				 "AllNumber="  & AllNumber & "," & _
				 "PackNumber=" & AllPack & "," & _
				 "PackWeight=" & PackWeight & "," & _
				 "Weight=" & Weight & "," & _
				 "Gross=" & Gross & "," & _
				 "Net=" & Net & "," & _
				 "Date1=" & "'" & Date1 & "'," & _
				 "Time1=" & "'" & Time1 & "'," & _
				 "Place1=" & "'" & Place1 & "'," & _
				 "Place2=" & "'" & Place2 & "'," & _
				 "Place3=" & "'" & Place3 & "'," & _
				 "RemainNumber=" & RemainNumber & "," & _
				 "RemainPack=" & RemainPack & "," & _
				 "RemainWeight=" & RemainWeight & ","  & _
				 "Operater=" & "'" & Operater1 & "'," & _
				 "Remark=" & "'" & Remark & "' " & _				 
				 "Where ContractID=" & "'" & ContractID & "'"

			'response.write sql
			 conn.Execute(sql) 
			 		         

	  
	  Case "Delete"      
	  		sql="Delete From OutStore Where RecNo=" & Request("RecNo")
	  		
	  		conn.Execute(sql)

	  
	  
	  Case "Return"
	  
   End Select

   if session("varClientID")="All" then
   			url="InStore_browse_all_frame.asp"
   else
   			url="InStore_browse_frame.asp"
   end if	
	
end if

'关闭数据库连接
conn.close
set conn=nothing

%>

<html><head><title>进仓管理</title><meta http-equiv=refresh content="0; url=<%=url%>">
</html>

