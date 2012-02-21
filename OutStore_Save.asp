<!-- #include file="data\config.asp" -->

<%

	sql="Update OutStore " & _
	    "set Gross=" & request("tGross") & ", " & _
	    "date1='" & trim(request("t3")) & "'," & _
		"date2='" & trim(request("t8")) & "' " & _	    
	    "where RecNo=" & request("RecNo")
	
	conn.Execute(sql)

	'关闭数据库连接
	conn.close
	set conn=nothing
	
    if Session("varClientID") = "All" then
        url="OutStore_Edit_All_Frame.asp"
    else
    	url="OutStore_Edit_Frame.asp"
    end if

%>

<html><head><title>进仓管理</title><meta http-equiv=refresh content="0; url=<%=url%>">
</html>