<!-- #include file="data\config.asp" -->
<table width="700">
<tr>

<td width="700">
<form name="my" method="POST" action="IOBrowseTable_ZhuangXieFei.asp" target="_blank">
<p style="margin-top: 4; margin-bottom: 4">
&nbsp;<select name="gsid" size=12>
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
 bYear=year(date())  '开始年份
 eYear=bYear         '结束年份
 bMonth=month(date())
 eMonth=bMonth
 bDay=1
 if isLeadYear(eYear) then
 	days(2)=29
 else
 	days(2)=28
 end if
 
 eDay=days(eMonth)
 
 %>
 <option value="All">显示全部</option> 
</select>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<p style="margin-top: 4; margin-bottom: 4">&nbsp; 按<select size="1" name="D1">   
  <option value="date1">日期</option>
  <option value="billno">单号</option>
  <option value="deliver">提货单位</option>
  <option value="contractid">合同号</option>
  <option value="goodsname">产品名</option>
  <option value="specs">规格</option>
</select>和<select size="1" name="D2">
  <option value="date1">日期</option>
  <option value="billno" selected>单号</option>
  <option value="deliver">提货单位</option>
  <option value="contractid">合同号</option>
  <option value="goodsname">产品名</option>
  <option value="specs">规格</option>
</select>排序   
</p>
<p style="margin-top: 4; margin-bottom: 4">筛选<select size="1" name="D3">
  <option value="contractid">合同号</option>
  <option value="goodsname">产品名</option>
  <option value="deliver">提货单位</option>
  <option value="specs">规格</option>
  <option value="billno">单号</option>
</select><select size="1" name="D4">
  <option selected value="--">--</option>
  <option value="=">等于</option>
  <option value="&gt;">大于</option>
  <option value="&gt;=">大于等于</option>
  <option value="&lt;">小于</option>
  <option value="&lt;=">小于等于</option>
  <option value="&lt;&gt;">不等于</option>
</select><input type="text" name="T1" size="9"> <select size="1" name="D5">   
  <option value="--">--</option>
  <option value="and">和</option>
  <option value="or">或</option>
</select><select size="1" name="D6">
  <option value="contractid">合同号</option>
  <option value="goodsname">产品名</option>
  <option value="deliver">提货单位</option>
  <option value="specs">规格</option>
  <option value="billno">单号</option>
</select><select size="1" name="D7">
  <option selected value="--">--</option>
  <option value="=">等于</option>
  <option value="&gt;">大于</option>
  <option value="&gt;=">大于等于</option>
  <option value="&lt;">小于</option>
  <option value="&lt;=">小于等于</option>
  <option value="&lt;&gt;">不等于</option>
</select><input type="text" name="T2" size="8">   
</p>
<p style="margin-top: 4; margin-bottom: 4"> &nbsp; 从<input type="text" name="tbyear" size="4" value=<% =bYear %>>年<input type="text" name="tbmonth" size="2" value=<% =bMonth %>>月      
<input type="text" name="tbday" size="2" value=<% =bDay %>>日 至 <input type="text" name="teyear" size="4" value=<% =eYear %>>年<input type="text" name="temonth" size="2" value=<% =bMonth %>>月<input type="text" name="teday" size="2" value=<% =eDay %>>日       

<br>
缷车费￥<input type="text" name="xiecheperton" size="4" value="11">元/吨(件) <br>
装车费￥<input type="text" name="zhuangcheperton" size="4" value="11">元/吨(件) <br>

</select><select size="1" name="D8">
  <option selected value="andunsuan">按吨算</option>
  <option value="anjiansuan">按件数算</option>
</select>


<input type="submit" value="计算装卸车费用" name="jy"> 

</p>

</form>  

</td>    


</tr>
</table>