<!-- #include file="data\config.asp" -->
<table width="700">
<tr>
<td width="700">
<form name="my" method="POST" action="reporty.asp" target="_blank">
<font size="4"><b>����ѡ��ͻ���</b></font>  
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
��<select name="ye">
<option><%=year(date())%></option>
<%
for i=2000 to year(date())
    response.write "<option>" & i & "</option>"
next
%>
</select>��&nbsp; <select name="mon"> 
<option><%=month(date())%></option>
<%
for i=1 to 12
    response.write "<option>" & i & "</option>"
next
%>
</select>�� 

<select name="ri"> 
<option>1</option>
<%
for i=1 to 31
    response.write "<option>" & i & "</option>"
next
%>
</select>��<br>
��<select name="ye1">
<option><%=year(date())%></option>
<%
for i=2000 to year(date())
    response.write "<option>" & i & "</option>"
next
%>
</select>��&nbsp; <select name="mon1"> 
<option><%=month(date())%></option>
<%
for i=1 to 12
    response.write "<option>" & i & "</option>"
next
%>
</select>�� 

<select name="ri1"> 
<option><%=day(date())%></option>
<%
for i=1 to 31
    response.write "<option>" & i & "</option>"
next
%>
</select>��
<input type="submit" value="����,�ֿ�,�����»��ܱ�" name="tj1">
</form>  
</td>    
</tr>
</table>