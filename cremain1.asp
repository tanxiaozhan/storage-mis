<!-- #include file="data\config.asp" -->
<table width="700">
<tr>

<td width="700">
<form name="my" method="POST" action="cremainty1.asp" target="_blank">
        <p>
          <input type="checkbox" name="lin" value="zero">
          显示结存为[0]的存货&nbsp;&nbsp; 到
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
          日</p>
        <p> 
          <input type="submit" value="查看结存" name="jy">
        </p>
      </form>  

</td>    


</tr>
</table>