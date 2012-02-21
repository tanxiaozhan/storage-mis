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
          response.write "旧密码输入错误！[<a href='javascript:history.back()'>返回</a>]"
          response.end
  else
          if request("newp")<>request("rnewp") then
              rs.close
              set rs=nothing
              conn.close
              set conn=nothing
              response.write "两次新密码输入不一致！[<a href='javascript:history.back()'>返回</a>]"
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
success="用户：" & request("us") & "-->密码修改成功！等待系统自动返回......"
url="index.asp"
%>

<html><head><title>用户改密</title><meta http-equiv=refresh content="2; url=<%=url%>">
<table width=100% height=100%>
<tr>
<td align=center>
<b><%=success%></b>
</td>
</tr>
</table>
</html>