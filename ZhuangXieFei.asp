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
 bYear=year(date())  '��ʼ���
 eYear=bYear         '�������
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
 <option value="All">��ʾȫ��</option> 
</select>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<p style="margin-top: 4; margin-bottom: 4">&nbsp; ��<select size="1" name="D1">   
  <option value="date1">����</option>
  <option value="billno">����</option>
  <option value="deliver">�����λ</option>
  <option value="contractid">��ͬ��</option>
  <option value="goodsname">��Ʒ��</option>
  <option value="specs">���</option>
</select>��<select size="1" name="D2">
  <option value="date1">����</option>
  <option value="billno" selected>����</option>
  <option value="deliver">�����λ</option>
  <option value="contractid">��ͬ��</option>
  <option value="goodsname">��Ʒ��</option>
  <option value="specs">���</option>
</select>����   
</p>
<p style="margin-top: 4; margin-bottom: 4">ɸѡ<select size="1" name="D3">
  <option value="contractid">��ͬ��</option>
  <option value="goodsname">��Ʒ��</option>
  <option value="deliver">�����λ</option>
  <option value="specs">���</option>
  <option value="billno">����</option>
</select><select size="1" name="D4">
  <option selected value="--">--</option>
  <option value="=">����</option>
  <option value="&gt;">����</option>
  <option value="&gt;=">���ڵ���</option>
  <option value="&lt;">С��</option>
  <option value="&lt;=">С�ڵ���</option>
  <option value="&lt;&gt;">������</option>
</select><input type="text" name="T1" size="9"> <select size="1" name="D5">   
  <option value="--">--</option>
  <option value="and">��</option>
  <option value="or">��</option>
</select><select size="1" name="D6">
  <option value="contractid">��ͬ��</option>
  <option value="goodsname">��Ʒ��</option>
  <option value="deliver">�����λ</option>
  <option value="specs">���</option>
  <option value="billno">����</option>
</select><select size="1" name="D7">
  <option selected value="--">--</option>
  <option value="=">����</option>
  <option value="&gt;">����</option>
  <option value="&gt;=">���ڵ���</option>
  <option value="&lt;">С��</option>
  <option value="&lt;=">С�ڵ���</option>
  <option value="&lt;&gt;">������</option>
</select><input type="text" name="T2" size="8">   
</p>
<p style="margin-top: 4; margin-bottom: 4"> &nbsp; ��<input type="text" name="tbyear" size="4" value=<% =bYear %>>��<input type="text" name="tbmonth" size="2" value=<% =bMonth %>>��      
<input type="text" name="tbday" size="2" value=<% =bDay %>>�� �� <input type="text" name="teyear" size="4" value=<% =eYear %>>��<input type="text" name="temonth" size="2" value=<% =bMonth %>>��<input type="text" name="teday" size="2" value=<% =eDay %>>��       

<br>
�����ѣ�<input type="text" name="xiecheperton" size="4" value="11">Ԫ/��(��) <br>
װ���ѣ�<input type="text" name="zhuangcheperton" size="4" value="11">Ԫ/��(��) <br>

</select><select size="1" name="D8">
  <option selected value="andunsuan">������</option>
  <option value="anjiansuan">��������</option>
</select>


<input type="submit" value="����װж������" name="jy"> 

</p>

</form>  

</td>    


</tr>
</table>