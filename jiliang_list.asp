<!-- #include file="data\config.asp" -->

<HTML>
<HEAD>
<TITLE>��˾��Ϣ</TITLE>
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<STYLE type=text/css>
<!--
A:link {text-decoration:none; color:"#0000ff"}
A:visited {text-decoration:none}
A:hover {text-decoration:underline; color:"#FF0033"}

body { font-family:"Verdana", "Arial", "Helvetica", "sans-serif"; font-size: 9pt}

p {  font-family:"Verdana", "Arial", "Helvetica", "sans-serif"; font-size: 9pt}

td {  font-family: "Verdana", "Arial", "Helvetica", "sans-serif"; font-size: 9pt}
-->

</STYLE>
</HEAD>
 <SCRIPT LANGUAGE ="VBSCRIPT">
  <!-- 
    Sub Checkdata()
      If trim(my.T2.Value) = Empty Then
        Window.Alert "����д����˾���ơ���"
        my.T2.focus()
        Exit Sub
      End If
      
       If trim(my.T3.Value) = Empty Then
        Window.Alert "����д����ַ����"
        my.T3.focus()
        Exit Sub
      End If

      if instr(1,trim(my.T2.Value),"'")>0 or instr(1,trim(my.T2.Value)," ")>0 then
          Window.Alert "����˾���ơ��в��ܺ��� ' �� '' ��"
          my.T1.focus()
          Exit Sub
      end if
     
       if instr(1,trim(my.T3.Value),"'")>0 or instr(1,trim(my.T3.Value)," ")>0 then
          Window.Alert "����ַ���в��ܺ��� ' �� '' ��"
          my.T3.focus()
          Exit Sub
      end if
      
       if instr(1,trim(my.T4.Value),"'")>0 or instr(1,trim(my.T4.Value)," ")>0 then
          Window.Alert "����ϵ�绰���в��ܺ��� ' �� '' ��"
          my.T4.focus()
          Exit Sub
      end if
      
      if trim(my.T5.Value)<>"" then  
        if instr(1,trim(my.T5.Value),"'")>0 or instr(1,trim(my.T5.Value)," ")>0 then
          Window.Alert "�����桱�в��ܺ��� ' �� '' ��"
          my.T5.focus()
          Exit Sub
        end if
      end if  
      
        
      my.Submit
    End Sub
    
  -->
  </SCRIPT>
  
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
  //window.alert a.length-1
  
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
        alert('����˾���ơ�����Ϊ�գ�');
	    document.my.T2.focus();
	    return false;
      }
    if((txt.indexOf("\"") != -1)||(txt.indexOf("\'") != -1)) 
      {
        alert('����˾���ơ��в��ܺ��� \' �� \" ��');
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
      
   
    return true;  
 }
-->
</script>

<BODY background=./image/blackground.gif bgColor=#ffffff>
<TABLE border=0 cellPadding=0 cellSpacing=0 height=55 width=839>
  <TR>
    <TD background=./image/top.gif colSpan=2 rowSpan=2 width=244>
      ��<DIV align=center><b><font face="��������" color="#FFFFFF" size="5"><span style="letter-spacing: 5pt">&nbsp;&nbsp;</span></font></b></DIV>
	</TD>
    <TD bgColor=#00b2eb height=1 width=595><IMG height=1 src="./image/pixel.gif" width=1></TD>
  </TR>
  <TR>
    <TD bgColor=#ffffff vAlign=bottom width="595">
      <DIV align=right>
      <p align="center">
      <img border="0" src="image/buttons.gif" width="460" height="40"></p>
      <TABLE align=right border=0 cellPadding=0 cellSpacing=0 height=32 width=312 vspace="0" hspace="0">
        <TR>
          <TD background=./image/button1.gif width=161>
            <DIV align=center><B><font size="4">�� ˾ �� Ϣ</font></B></DIV>                    
		  </TD> 
          <TD vAlign=bottom width="147"> 
              <DIV align=right></DIV> 
		  </TD> 
		</TR> 
	  </TABLE> 
	  </DIV> 
	</TD> 
  </TR> 
</TABLE> 
<TABLE border=0 cellPadding=0 cellSpacing=0 width=844 height="327"> 
  <TR> 
    <TD height=1 width=248> 
    <p align="center"></TD> 
    <TD rowSpan=2 width=8 height="327">��</TD> 
    <TD height=1 width=587></TD> 
  </TR> 
  <TR> 
    <TD vAlign=top width=248 height="321"> 
      <TABLE align=right border=0 cellPadding=4 cellSpacing=4 height=270 width=141> 
        <TR bgColor=#29a8cd> 
          <TD> 
            <DIV align=center><a href="index.asp?op=reset"><b><font size="2" color="#FFFFFF">���µ�¼</font></b></a></DIV> 
		  </TD> 
		</TR> 
        <TR bgColor=#2cb1bc> 
          <TD> 
            <DIV align=center><b><A href="instore.asp"><font color="#ffffff" size="2"> 
              ���ֹ���</font></A></b></DIV> 
		  </TD> 
		</TR> 
        <TR bgColor=#35a8a5> 
          <TD> 
            <DIV align=center><b><A href="outstore.asp"><font color="#ffffff" size="2"> 
              ���ֹ���</font></A></b></DIV> 
		  </TD> 
		</TR> 
        <TR bgColor=#35a88e> 
          <TD> 
            <DIV align=center><b> 
              <a href="Client_List.htm"> 
              <font color="#ffffff" size="2"> 
              �ͻ���Ϣ</font></a></b></DIV> 
		  </TD> 
		</TR> 
        <TR bgColor=#37a67c> 
          <TD> 
            <DIV align=center><b><A href="user.htm"><font color="#ffffff" size="2"> 
              �û�</font><font size="2" color="#FFFFFF">����</font></A></b></DIV> 
		  </TD> 
		</TR> 
        <TR bgColor=#339966> 
          <TD> 
            <DIV align=center><b><A href="jiliang_list.htm"><font color="#ffffff" size="2"> 
              ������λ</font></A></b></DIV> 
		  </TD> 
		</TR> 
        <TR bgColor=#339966> 
          <TD> 
            <p align="center"><A href="report.htm"><font size="2" color="#FFFFFF"><b>������Ʒ�</b></font></a></TD>     
		</TR>     
        <TR bgColor=#339966> 
          <TD> 
            <p align="center"><a href="cremain.htm"><font size="2" color="#FFFFFF"><b>������</b></font></a></TD>     
		</TR>     
        <TR bgColor=#339966> 
          <TD> 
            <p align="center"><A href="recover.htm"><font size="2" color="#FFFFFF"><b>���ݻָ�</b></font></a></TD>     
		</TR>     
        <TR bgColor=#339966> 
          <TD> 
            <DIV align=center><b><A href="support.htm"><font color="#ffffff" size="2"> 
              ��&nbsp; ��</font></A></b></DIV>                                
		  </TD>     
		</TR>     
        </TABLE>     
      <P align=right>��</P>
	</TD>
<%

    Set rs =server.createobject("adodb.recordset")

	sql="select * from Master"
	rs.open sql,conn,3,3
	if rs.recordcount<1 then
		rs.addnew
	end if
			
%>

    <TD width="587" valign="top" height="321"> 
      <form method="POST" name="my" action="jiliang_add.asp" onsubmit="javascript:return validate_form();">
        <p style="line-height: 100%; margin-top: 5; margin-bottom: 1">��</p>
        <p style="line-height: 100%; margin-top: 5; margin-bottom: 1"><font face="����" size="3"><b>
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ��˾���ƣ�<input type="text" name="T2" size="55" maxlength="50" value="<% = rs("name") %>"> *</b></font></p>                    
        <p style="line-height: 100%; margin-top: 5; margin-bottom: 1"><font face="����" size="3"><b>
        ����&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ��ַ��<input type="text" name="T3" size="55" maxlength="50" value="<% = rs("Address") %>">            
        *</b></font></p>                    
        <p style="line-height: 100%; margin-top: 5; margin-bottom: 1"><font face="����" size="3"><b>
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ��ϵ�绰��<input type="text" name="T4" size="55" maxlength="50" value="<% = rs("Tel") %>" >            
        &nbsp;</b></font></p>                   
        <p style="line-height: 100%; margin-top: 5; margin-bottom: 1"><font face="����" size="3"><b>
        ����&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ���棺<input type="text" name="T5" size="55" maxlength="50" value="<% = rs("Fax") %>" ></b></font></p>          
        <p style="line-height: 100%; margin-top: 5; margin-bottom: 1">
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;         
        <input type="hidden" name="khtj" value="tj">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</p>
        <p style="line-height: 100%; margin-top: 5; margin-bottom: 1">
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <font size="3" face="����"><b><input type="submit" value="����" name="tj" style="font-family: ����; font-size: 12pt; font-weight: bold"></b></font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
        <input type="reset" value="����" name="cz" OnClick="vbscript:history.back()" style="font-family: ����; font-size: 12pt; font-weight: bold">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</p>
        <p style=" 150%; margin-top: 0; margin-bottom: 0" line-height:>��</p>
        <p style="line-height: 150%; margin-top: 0; margin-bottom: 0" align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
        ��ע:*Ϊ���ֶΡ�</p>
      </form>
<%
  rs.close()
%>
      <p>��<p align="left">��</td>
  </tr>
  </table>
</body>
</html>