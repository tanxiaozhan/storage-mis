<!-- #include file="data\config.asp" -->

<%
	set rs=server.createobject("adodb.recordset")
	sql="select * from OutStore where RecNo=" & Request("RecNo")
	rs.open sql,conn,1,1
	if rs.eof then
		rs.close
		set rs=nothing
		if session("varClientID")="All" then
			response.redirect "OutStore_Edit_all_frame.asp"	
		else
			response.redirect "OutStore_Edit_frame.asp"	
		end if	
		response.end
	end if
	
	vInStoreRecNo=cstr(rs("InStoreRecNo"))

   '删除出仓记录之前，恢复进仓结余数据
	sql="Update InStore " & _
	    "set RemainNumber=RemainNumber + " & rs("Number1") & "," & _
	    "RemainPack= RemainPack + " & rs("Pack1") & "," & _
	    "RemainWeight=RemainWeight + " & rs("Weight") & " " & _
	    "where RecNo=" & rs("InStoreRecNo")
	
	rs.close
	'response.write (sql)
	'response.end
	conn.Execute(sql)

	sql="Select * from InStore where RecNo=" & vInStoreRecNo
	rs.open sql,conn,3,3
	if rs("RemainNumber")>rs("AllNumber") then
		rs("RemainNumber")=rs("AllNumber")
		rs.update
	end if
	if rs("RemainPack") > rs("PackNumber") then	
		rs("RemainPack")= rs("PackNumber")
		rs.update
	end if
	if cdbl(rs("RemainWeight")) > cdbl(rs("Weight")) then
		rs("RemainWeight") = rs("weight")
		rs.update
	end if
	
	rs.close
	set rs=nothing
	

	'删除前将其备份
	sql="insert into DelOutStore select * from OutStore where RecNo=" & Request("RecNo")
	conn.Execute(sql)
	
	'删除出仓记录
	sql="delete from OutStore where RecNo=" & Request("RecNo")
	'response.write sql
	
	conn.Execute(sql)
    
    if Session("varClientID") = "All" then
        url="OutStore_Edit_All_Frame.asp"
    else
    	url="OutStore_Edit_Frame.asp"
    end if


	'关闭数据库连接
	conn.close
	set conn=nothing

%>

<html><head><title>进仓管理</title><meta http-equiv=refresh content="0; url=<%=url%>">
</html>

