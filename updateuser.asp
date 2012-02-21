<%
if session("userright")<>"仓库主管" and session("userright")<>"仓管" then
   response.write "你没有使用的权限！"
   response.end
end if
%>
<!-- #include file="data\config.asp" -->
<%
if request("T1")<>trim(request("T2")) then
    Set rs =server.createobject("adodb.recordset")
    sql="select * from ClientInfo where ClientName='" & trim(request("T2")) & "'"
    rs.Open sql,conn,3,3
    if rs.recordcount>0 then
       rs.close
       conn.close
       set rs=nothing
       set conn=nothing
       
       response.write "客户名称已经存在，请重新输入！[<a href='javascript:history.back()'>返回</a>]"
       response.end
 
    end if
    rs.close
    set rs=nothing
    
    if trim(request("T5"))<>"" then
       sql="update ClientInfo set ClientName='" & trim(request("T2"))  & "',address='" & trim(request("T3")) & "',Telphone='" & trim(request("T4")) & "',Fax='" & trim(request("T5")) & "',LinkMan='" & trim(request("T6")) & "' where id=" & request("idu") 
    else
       sql="update ClientInfo set ClientName='" & trim(request("T2"))  & "',address='" & trim(request("T3")) & "',Telphone='" & trim(request("T4")) & "',Fax='" & space(2) & "',LinkMan='" & trim(request("T6")) & "' where id=" & request("idu")         
    end if
    
else
    if trim(request("T5"))<>"" then
       sql="update ClientInfo set address='" & trim(request("T3")) & "',Telphone='" & trim(request("T4")) & "',Fax='" & trim(request("T5")) & "',LinkMan='" & trim(request("T6")) & "',price=" & trim(request("T7")) & " where id=" & request("idu") 
    else
       sql="update ClientInfo set address='" & trim(request("T3")) & "',Telphone='" & trim(request("T4")) & "',Fax='" & space(2) & "',LinkMan='" & trim(request("T6")) & "',price=" & trim(request("T7")) & " where id=" & request("idu")
    end if   
end if  

conn.execute(sql)

conn.close
set conn=nothing

success="客户信息修改成功！等待系统自动返回......"
url="client_list.asp"
%>

<html><head><title>修改客户信息</title><meta http-equiv=refresh content="2; url=<%=url%>">
<table width=100% height=100%>
<tr>
<td align=center>
<b><%=success%></b>
</td>
</tr>
</table>
</html>