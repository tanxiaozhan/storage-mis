<%
if session("userright")<>"�ֿ�����" and session("userright")<>"�ֹ�" then
   response.write "��û��ʹ�õ�Ȩ�ޣ�"
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
       
       response.write "�ͻ������Ѿ����ڣ����������룡[<a href='javascript:history.back()'>����</a>]"
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

success="�ͻ���Ϣ�޸ĳɹ����ȴ�ϵͳ�Զ�����......"
url="client_list.asp"
%>

<html><head><title>�޸Ŀͻ���Ϣ</title><meta http-equiv=refresh content="2; url=<%=url%>">
<table width=100% height=100%>
<tr>
<td align=center>
<b><%=success%></b>
</td>
</tr>
</table>
</html>