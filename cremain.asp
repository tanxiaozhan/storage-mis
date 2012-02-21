<!-- #include file="data\config.asp" -->
<table width="700">
<tr>
<td width="700">
<form name="my" method="POST" action="cremainty.asp" target="_blank">
        <p>
          <select name="gsid" size=12>
            <%
  Set rs =server.createobject("adodb.recordset")
  sql="select distinct ClientName,ID from ClientInfo"
  rs.open sql,conn,3,3
  
  while not rs.eof
      response.write "<option value='" & rs("ID") & "'>" & rs("ClientName") & "</option>"
      rs.movenext
 wend
 rs.close
 set rs=nothing
 set conn=nothing
 %>
          </select>
          <input type="checkbox" name="lin" value="zero">
          显示结存为[0]的存货&nbsp;&nbsp; </p>
        <p>到
          <select name="ye">
            <option><%=year(date())%></option>
            <%
for i=2000 to year(date())
    response.write "<option>" & i & "</option>"
next
%>
          </select>
          年&nbsp; 
          <select name="mon">
            <option><%=month(date())%></option>
            <%
for i=1 to 12
    response.write "<option>" & i & "</option>"
next
%>
          </select>
          月 
          <select name="ri">
            <option><%=day(date())%></option>
            <%
for i=1 to 31
    response.write "<option>" & i & "</option>"
next
%>
          </select>
          日 
          <input type="submit" value="查看结存" name="jy">
        </p>
      </form>  

</td>    


</tr>
</table>