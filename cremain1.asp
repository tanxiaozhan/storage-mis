<!-- #include file="data\config.asp" -->
<table width="700">
<tr>

<td width="700">
<form name="my" method="POST" action="cremainty1.asp" target="_blank">
        <p>
          <input type="checkbox" name="lin" value="zero">
          ��ʾ���Ϊ[0]�Ĵ��&nbsp;&nbsp; ��
          <select name="ye">
            <option><%=year(date())%></option>
            <%
for i=2000 to year(date())
    response.write "<option>" & i & "</option>"
next
%>
          </select>
          ��&nbsp; 
          <select name="mon">
            <option><%=month(date())%></option>
            <%
for i=1 to 12
    response.write "<option>" & i & "</option>"
next
%>
          </select>
          �� 
          <select name="ri">
            <option><%=day(date())%></option>
            <%
for i=1 to 31
    response.write "<option>" & i & "</option>"
next
%>
          </select>
          ��</p>
        <p> 
          <input type="submit" value="�鿴���" name="jy">
        </p>
      </form>  

</td>    


</tr>
</table>