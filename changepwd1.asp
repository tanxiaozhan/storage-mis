<%
function jiami(strf)
   
   randomize
   tmp=""
   
   sz=(rnd * 25) + 90'65
   for i=1 to len(strf)
       tmp=tmp & chr((asc(mid(strf,i,1)) xor sz))
   next
   
   jiami=tmp & chr(sz)
end function


function jiemi(strf)
   
   sz=asc(right(strf,1))
   
   tmp=""
   
   for i=1 to len(strf)-1
       tmp=tmp & chr((asc(mid(strf,i,1)) xor sz))
   next
   
   jiemi=tmp
      
end function

%>
<%
    dim conn
    dim  connectstr
    set  conn=server.createobject("adodb.connection")
    'pathstr=split(server.mappath("chat.mdb"),"\")
    'pathstr(0) & "\" pathstr(1) & "data\" & "chat.mdb"
    connectstr="provider=microsoft.jet.oledb.4.0; jet oledb:database password=office;data source=" & server.mappath("data\SMIS.mdb")
    'Response.Write connectstr
    'Response.End 
    conn.Open connectstr 
 %>
<%
  sql="select * from userInfo where user='" & request("us") &"'"
  
  Set rs =server.createobject("adodb.recordset")
  rs.Open sql,conn,3,3
  

 if request("oldp")<>trim(jiemi(rs("pwd"))) then
          rs.close
          set rs=nothing
          conn.close
          set conn=nothing
          response.write "�������������[<a href='javascript:history.back()'>����</a>]"
          response.end
  else
          if request("newp")<>request("rnewp") then
              rs.close
              set rs=nothing
              conn.close
              set conn=nothing
              response.write "�������������벻һ�£�[<a href='javascript:history.back()'>����</a>]"
              response.end
           else
              
              if trim(request("newp"))="" then
                 rs("pwd")=jiami(space(2))
              else
                 rs("pwd")=jiami(trim(request("newp"))) 
              end if
              
              rs.update
          end if
  end if         

  
  
rs.close
set rs=nothing  
conn.close
set conn=nothing
success="�û���" & request("us") & "-->�����޸ĳɹ����ȴ�ϵͳ�Զ�����......"
url="index.asp"
%>

<html><head><title>�û�����</title><meta http-equiv=refresh content="2; url=<%=url%>">
<table width=100% height=100%>
<tr>
<td align=center>
<b><%=success%></b>
</td>
</tr>
</table>
</html>