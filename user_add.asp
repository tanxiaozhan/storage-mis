<%
if session("userright")<>"�ֿ�����" then
   response.write "��û��ʹ�õ�Ȩ�ޣ���<a href='javascript:history.back()'>����</a>��"
   response.end
end if
%>
<!-- #include file="data\config.asp" -->
<%

if request("tj")="�ύ" then
  
  sql="select * from userInfo where user='" & trim(request("T2")) &"'"
  
  Set rs =server.createobject("adodb.recordset")
  rs.Open sql,conn,3,3
 
  if rs.recordcount>0 then 
     rs.close
     set rs=nothing
     conn.close
     set conn=nothing
     
     response.write "�û������Ѿ����ڣ����������룡[<a href='javascript:history.back()'>����</a>]"
     response.end
  end if
  
  zd=""
  for i=0 to rs.fields.count-1
      if i<>rs.fields.count-1 then
          zd=zd & rs.fields(i).name & ","
      else
          zd=zd & rs.fields(i).name
      end if   
  next 
  
  'rs.close
  'set rs=nothing
  
  'if trim(request("T3"))="" then
     'sql="insert into ClientInfo (ClientName,Address,Telphone,Fax,LinkMan) values('" & trim(request("T2")) & "','" & trim(request("T3")) & "','" & trim(request("T4")) & "','" & space(2) & "','" & trim(request("T6")) & "')"
     sql="insert into userInfo (" & zd & ") values('" & trim(request("T2")) & "','��','" & trim(request("T4")) & "')"
  'else 
     sql="insert into userInfo (" & zd & ") values('" & trim(request("T2")) & "','" & trim(request("T3")) & "','" & trim(request("T4")) & "')"          
  'end if
  
  'response.write sql    
  'conn.Execute(sql)
  'response.end   
  
  rs.addnew
  rs("user")=trim(request("T2"))
  if trim(request("T3"))="" then
      rs("pwd")=jiami(space(2))
  else
      rs("pwd")=jiami(trim(request("T3")))
  end if
  
  rs("Right")=trim(request("T4"))
  rs.update
  
  rs.close
  set rs=nothing  
  
end if


conn.close
set conn=nothing

'response.write "�û���ӳɹ����ȴ�ϵͳ�Զ�����......"
'response.redirect "client_add.htm"

success="�û���ӳɹ����ȴ�ϵͳ�Զ�����......"
url="user.htm"

%>
<html><head><title>����û�</title><meta http-equiv=refresh content="1; url=<%=url%>">
<table width=100% height=100%>
<tr>
<td align=center>
<b><%=success%></b>
</td>
</tr>
</table>
</html>