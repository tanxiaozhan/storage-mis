<%
  if session("username")="" or session("userright")="" then 
     response.write "你还没有登录系统！请【</b></font><a href='index.asp'>登录</a>】"
     response.end
  end if
%>
<HTML>
<HEAD>
<TITLE>仓库管理系统</TITLE>
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
        Window.Alert "请填写“客户名称”！"
        my.T2.focus()
        Exit Sub
      End If
      
       If trim(my.T3.Value) = Empty Then
        Window.Alert "请填写“地址”！"
        my.T3.focus()
        Exit Sub
      End If

      If trim(my.T4.Value) = Empty Then
        Window.Alert "请填写“联系电话”！"
        my.T4.focus()
        Exit Sub
      End If

      If trim(my.T6.Value) = Empty Then
        Window.Alert "请填写“联系人”！"
        my.T6.focus()
        Exit Sub
      End If
      
      if instr(1,trim(my.T2.Value),"'")>0 or instr(1,trim(my.T2.Value)," ")>0 then
          Window.Alert "“客户名称”中不能含有 ' 和 '' ！"
          my.T1.focus()
          Exit Sub
      end if
     
       if instr(1,trim(my.T3.Value),"'")>0 or instr(1,trim(my.T3.Value)," ")>0 then
          Window.Alert "“客户名称”中不能含有 ' 和 '' ！"
          my.T3.focus()
          Exit Sub
      end if
      
       if instr(1,trim(my.T4.Value),"'")>0 or instr(1,trim(my.T4.Value)," ")>0 then
          Window.Alert "“客户名称”中不能含有 ' 和 '' ！"
          my.T4.focus()
          Exit Sub
      end if
      
      if trim(my.T5.Value)<>"" then  
        if instr(1,trim(my.T5.Value),"'")>0 or instr(1,trim(my.T5.Value)," ")>0 then
          Window.Alert "“客户名称”中不能含有 ' 和 '' ！"
          my.T5.focus()
          Exit Sub
        end if
      end if  
      
       if instr(1,trim(my.T6.Value),"'")>0 or instr(1,trim(my.T6.Value)," ")>0 then
          Window.Alert "“客户名称”中不能含有 ' 和 '' ！"
          my.T6.focus()
          Exit Sub
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
        alert('“客户名称”不能为空！');
	    document.my.T2.focus();
	    return false;
      }
    if((txt.indexOf("\"") != -1)||(txt.indexOf("\'") != -1)) 
      {
        alert('“客户名称”中不能含有 \' 和 \" ！');
        document.my.T2.focus();
	    return false;
      }
  
    txt=document.my.T3.value;
    if ((txt=="")||(txt==sp(txt.length)))
      {
        alert('“地址”不能为空！');
	    document.my.T3.focus();
	    return false;
      }
    if((txt.indexOf("\"") != -1)||(txt.indexOf("\'") != -1)) 
      {
        alert ('“地址”中不能含有 \' 和 \" ！');
        document.my.T3.focus();
	    return false;
      }  
      
    txt=document.my.T4.value;
    if ((txt=="")||(txt==sp(txt.length)))
      {
        alert('“联系电话”不能为空！');
	    document.my.T4.focus();
	    return false;
      }
    if((txt.indexOf("\"") != -1)||(txt.indexOf("\'") != -1)) 
      {
        alert ('“联系电话”中不能含有 \' 和 \" ！');
        document.my.T4.focus();
	    return false;
      }      

    txt=document.my.T5.value;
    if((txt.indexOf("\"") != -1)||(txt.indexOf("\'") != -1)) 
      {
        alert ('“传真”中不能含有 \' 和 \" ！"');
        document.my.T5.focus();
	    return false;
      }
      
    txt=document.my.T6.value;
    if ((txt=="")||(txt==sp(txt.length)))
      {
        alert('“联系人”不能为空！');
	    document.my.T6.focus();
	    return false;
      }      
    if((txt.indexOf("\"") != -1)||(txt.indexOf("\'") != -1)) 
      {
        alert ('“联系人”中不能含有 \' 和 \" ！');
        document.my.T6.focus();
	    return false;
      }    
   
    txt=document.my.T7.value;
    if ((txt=="")||(txt==sp(txt.length)))
      {
        alert('“单价”不能为空！');
	    document.my.T7.focus();
	    return false;
      }      
    if(!checknum(txt)) 
      {
        alert ('“单价”必须为数字 ！');
        document.my.T7.focus();
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
      　<DIV align=center><b><font face="华文隶书" color="#FFFFFF" size="5"><span style="letter-spacing: 5pt">&nbsp;&nbsp;</span></font></b></DIV>
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
            <DIV align=center><font size="4"><b>登 录 成 功</b></font></DIV>          
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
    <TD rowSpan=2 width=8 height="327">　</TD> 
    <TD height=1 width=587></TD> 
  </TR> 
  <TR> 
    <TD vAlign=top width=248 height="321"> 
      <TABLE align=right border=0 cellPadding=4 cellSpacing=4 height=270 width=141> 
        <TR bgColor=#29a8cd> 
          <TD> 
            <DIV align=center><a href="index.asp?op=reset"><b><font size="2" color="#FFFFFF">重新登录</font></b></a></DIV> 
		  </TD> 
		</TR> 
        <TR bgColor=#2cb1bc> 
          <TD> 
            <DIV align=center><b><A href="instore.asp"><font color="#ffffff" size="2"> 
              进仓管理</font></A></b></DIV> 
		  </TD> 
		</TR> 
        <TR bgColor=#35a8a5> 
          <TD> 
            <DIV align=center><b><A href="outstore.asp"><font color="#ffffff" size="2"> 
              出仓管理</font></A></b></DIV> 
		  </TD> 
		</TR> 
        <TR bgColor=#35a88e> 
          <TD> 
            <DIV align=center><b><A href="client_list.htm"><font color="#ffffff" size="2"> 
              客户信息</font></A></b></DIV> 
		  </TD> 
		</TR> 
        <TR bgColor=#37a67c> 
          <TD> 
            <DIV align=center><b><A href="user.htm"><font color="#ffffff" size="2"> 
              用户</font><font size="2" color="#FFFFFF">管理</font></A></b></DIV> 
		  </TD> 
		</TR> 
        <TR bgColor=#339966> 
          <TD> 
            <DIV align=center><b><a href="jiliang_list.htm"><font size="2" color="#FFFFFF">选 
              项</font></a></b></DIV> 
		  </TD> 
		</TR> 
        <TR bgColor=#339966> 
          <TD> 
            <p align="center"><A href="report.htm"><font size="2" color="#FFFFFF"><b>报表与计费</b></font></a></TD>     
		</TR>     
        <TR bgColor=#339966> 
          <TD> 
            <p align="center"><a href="cremain.htm"><font size="2" color="#FFFFFF"><b>结存情况</b></font></a></TD>     
		</TR>     
        <TR bgColor=#339966> 
          <TD> 
            <p align="center"><A href="recover.htm"><font size="2" color="#FFFFFF"><b>数据恢复</b></font></a></TD>     
		</TR>     
        <TR bgColor=#339966> 
          <TD> 
            <DIV align=center><b><A href="support.htm"><font color="#ffffff" size="2"> 
              帮&nbsp; 助</font></A></b></DIV>                      
		  </TD>     
		</TR>     
        </TABLE>     
      <P align=right>　</P>
	</TD>
    <TD width="587" valign="top" height="321"> 
      <p align="center">
      <img border="0" src="image/lg.jpg" width="550" height="350"><p align="left">　</td>
  </tr>
  </table>
</body>
</html>