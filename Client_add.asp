<!-- #include file="data\config.asp" -->
<%

if request("khtj")="tj" then
  
  sql="select * from ClientInfo where ClientName='" & trim(request("T2")) &"'"
  
  Set rs =server.createobject("adodb.recordset")
  rs.Open sql,conn,3,3
 
  if rs.recordcount>0 then 
     rs.close
     set rs=nothing
     conn.close
     set conn=nothing
     
     response.write "�ͻ������Ѿ����ڣ����������룡[<a href='javascript:history.back()'>����</a>]"
     response.end
  end if
  
  zd=""
  for i=1 to rs.fields.count-1
      if i<>rs.fields.count-1 then
          zd=zd & rs.fields(i).name & ","
      else
          zd=zd & rs.fields(i).name
      end if   
  next 
  
  rs.close
  set rs=nothing
  
  if trim(request("T5"))="" then
     'sql="insert into ClientInfo (ClientName,Address,Telphone,Fax,LinkMan) values('" & trim(request("T2")) & "','" & trim(request("T3")) & "','" & trim(request("T4")) & "','" & space(2) & "','" & trim(request("T6")) & "')"
     sql="insert into ClientInfo (" & zd & ") values('" & trim(request("T2")) & "','" & trim(request("T3")) & "','" & trim(request("T4")) & "','" & space(2) & "','" & trim(request("T6")) & "'," & trim(request("T7")) &")"
  else 
     sql="insert into ClientInfo (" & zd & ") values('" & trim(request("T2")) & "','" & trim(request("T3")) & "','" & trim(request("T4")) & "','" & trim(request("T5")) & "','" & trim(request("T6")) & "'," & trim(request("T7")) &")"         
  end if
  
  'response.write sql    
  conn.Execute(sql)
     
end if

conn.close
set conn=nothing

'response.write "�ͻ���Ϣ���ӳɹ����ȴ�ϵͳ�Զ�����......"
'response.redirect "client_add.htm"

success="�ͻ���Ϣ���ӳɹ����ȴ�ϵͳ�Զ�����......"
url="client_add.htm"

%>
<html><head><title>�ͻ���Ϣ</title><meta http-equiv=refresh content="1; url=<%=url%>">
<table width=100% height=100%>
<tr>
<td align=center>
<b><%=success%></b>
</td>
</tr>
</table>
</html>