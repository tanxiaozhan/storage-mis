<!-- #include file="data\config.asp" -->
<script language="javascript">
<!--
function sp(a)
{
  var i=1;
  
  var k="";
  
  for(i=1;i<=a;i++)
    k=k+" ";
    
  return k;  
}

function checknum(a)
{ 
  for(i=0;i<=a.length-1;i++)
    {
      if(((a.substring(i,i+1)>'.')||(a.substring(i,i+1)<'.'))&&((a.substring(i,i+1)>' ')||(a.substring(i,i+1)<' ')))
       {
         if((a.substring(i,i+1)<'0')||(a.substring(i,i+1)>'9'))             
            return false;
       }      
    }
    
  return true;  
}

function validate_form()
{
    var txt;
    txt=document.my.T2.value;
    if ((txt=="")||(txt==sp(txt.length)))
      {
        alert('���ͻ����ơ�����Ϊ�գ�');
	    document.my.T2.focus();
	    return false;
      }
    if((txt.indexOf("\"") != -1)||(txt.indexOf("\'") != -1)) 
      {
        alert('���ͻ����ơ��в��ܺ��� \' �� \" ��');
        document.my.T2.focus();
	    return false;
      }
  
    txt=document.my.T3.value;
    if ((txt=="")||(txt==sp(txt.length)))
      {
        alert('����ַ������Ϊ�գ�');
	    document.my.T3.focus();
	    return false;
      }
    if((txt.indexOf("\"") != -1)||(txt.indexOf("\'") != -1)) 
      {
        alert ('����ַ���в��ܺ��� \' �� \" ��');
        document.my.T3.focus();
	    return false;
      }  
      
    txt=document.my.T4.value;
    if ((txt=="")||(txt==sp(txt.length)))
      {
        alert('����ϵ�绰������Ϊ�գ�');
	    document.my.T4.focus();
	    return false;
      }
    if((txt.indexOf("\"") != -1)||(txt.indexOf("\'") != -1)) 
      {
        alert ('����ϵ�绰���в��ܺ��� \' �� \" ��');
        document.my.T4.focus();
	    return false;
      }      

    txt=document.my.T5.value;
    if((txt.indexOf("\"") != -1)||(txt.indexOf("\'") != -1)) 
      {
        alert ('�����桱�в��ܺ��� \' �� \" ��"');
        document.my.T5.focus();
	    return false;
      }
      
    txt=document.my.T6.value;
    if ((txt=="")||(txt==sp(txt.length)))
      {
        alert('����ϵ�ˡ�����Ϊ�գ�');
	    document.my.T6.focus();
	    return false;
      }      
    if((txt.indexOf("\"") != -1)||(txt.indexOf("\'") != -1)) 
      {
        alert ('����ϵ�ˡ��в��ܺ��� \' �� \" ��');
        document.my.T6.focus();
	    return false;
      }    
      
    txt=document.my.T7.value;
    if ((txt=="")||(txt==sp(txt.length)))
      {
        alert('�����ۡ�����Ϊ�գ�');
	    document.my.T7.focus();
	    return false;
      }      
    if(!checknum(txt)) 
      {
        alert ('�����ۡ�����Ϊ���� ��');
        document.my.T7.focus();
	    return false;
      }   
  
      
    return true;  
 }
-->
</script>
<%
  Set rs =server.createobject("adodb.recordset")
  sql="select * from Clientinfo where id=" & request("id")
  rs.Open sql,conn,3,3 
  if rs("price")<1  and rs("price")>0 then
      pr="0" & rs("price")
  else
      pr=rs("price")
  end if       

%>
<form name="my" action="updateuser.asp" onsubmit="javascript:return validate_form();">       
         <input type="hidden" name="idu" value=<%=rs("id")%>>
         <input type="hidden" name="T1" value=<%=rs("ClientName")%>>
         <p style="line-height: 100%; margin-top: 5; margin-bottom: 1"><font face="����" size="3"><b>
        &nbsp;&nbsp;&nbsp;&nbsp;
        �ͻ����ƣ�<input type="text" name="T2" size="55" maxlength="50" value=<%=rs("ClientName")%>>   
        *</b> </font></p> 
        <p style="line-height: 100%; margin-top: 5; margin-bottom: 1"><font face="����" size="3"><b>
        &nbsp;&nbsp;&nbsp;&nbsp;
        ������ַ��<input type="text" name="T3" size="55" maxlength="50" value=<%=rs("address")%>> 
        *</b></font></p>   
        <p style="line-height: 100%; margin-top: 5; margin-bottom: 1"><font face="����" size="3"><b>
        &nbsp;&nbsp;&nbsp;&nbsp;
        ��ϵ�绰��<input type="text" name="T4" size="55" maxlength="50" value=<%=rs("Telphone")%>> 
        *</b></font></p>   
        <p style="line-height: 100%; margin-top: 5; margin-bottom: 1"><font face="����" size="3"><b>
        ����&nbsp;&nbsp;&nbsp;&nbsp; ���棺<input type="text" name="T5" size="55" maxlength="50" value=<%=rs("fax")%>></b></font></p>
        <p style="line-height: 100%; margin-top: 5; margin-bottom: 1"><font face="����" size="3"><b>
        ��&nbsp;&nbsp;&nbsp;&nbsp; ��ϵ�ˣ�<input type="text" name="T6" size="55" maxlength="50" value=<%=rs("linkman")%>> 
        *</b></font></p>   
        <p style="line-height: 100%; margin-top: 5; margin-bottom: 1"><b>
        <font face="����">&nbsp; </font></b><font face="����" size="3"><b>
        ����(��/Ԫ)��<input type="text" name="T7" size="55" maxlength="50" value=<%=pr%>> 
        *</b></font></p>   
        <p style="line-height: 100%; margin-top: 5; margin-bottom: 1"></p>
<p>&nbsp; 
        <input type="reset" value="����">&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp; 
<input type="submit" name="gai" value="�޸�">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp; 
<input type="button" value="ɾ��" onclick="javascript:location.href='deluser.asp?id='+document.my.idu.value">&nbsp;&nbsp;&nbsp;&nbsp; 
&nbsp;&nbsp;&nbsp; <input type="button" value="����" onclick="vbscript:history.back()"> 
</form>   
<%   
rs.close
set rs=nothing
conn.close
set conn=nothing
%>

</p>