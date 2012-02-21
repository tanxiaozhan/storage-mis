<!-- #include file="data\config.asp" -->
<table width="700">
<tr>
<td width="700">
<form name="my" method="POST" action="reporty.asp" target="_blank">
<font size="4"><b>【请选择客户】</b></font>  
<br>
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
<br> 
从<select name="ye">
<option><%=year(date())%></option>
<%
for i=2000 to year(date())
    response.write "<option>" & i & "</option>"
next
%>
</select>年&nbsp; <select name="mon"> 
<option><%=month(date())%></option>
<%
for i=1 to 12
    response.write "<option>" & i & "</option>"
next
%>
</select>月 

<select name="ri"> 
<option>1</option>
<%
for i=1 to 31
    response.write "<option>" & i & "</option>"
next
%>
</select>日<br>
到<select name="ye1">
<option><%=year(date())%></option>
<%
for i=2000 to year(date())
    response.write "<option>" & i & "</option>"
next
%>
</select>年&nbsp; <select name="mon1"> 
<option><%=month(date())%></option>
<%
for i=1 to 12
    response.write "<option>" & i & "</option>"
next
%>
</select>月 

<select name="ri1"> 
<option><%=day(date())%></option>
<%
for i=1 to 31
    response.write "<option>" & i & "</option>"
next
%>
</select>日
<input type="submit" value="挑拨,仓库,仓租月汇总表" name="tj1">
</form>  
</td>    
</tr>
</table>