<!-- #include file="data\config.asp" -->

<HTML>
<HEAD>
<TITLE>出仓管理</TITLE>
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

<%
' if session("username")="区福全" then 
iwidth="1140"
 ' else
 '  iwidth="700"
 '  end if



%>


<%
	Session("varClientID")=Request.form("ClientName")
	Session("varBillNo")=Request.form("tBillNo")
	Session("varContractID")=Request.form("tContractID")
	Session("varSpecs")=Request.form("tSpecs")
    Session("varGoodsName")=Request.form("tGoodsName")

%>
 <SCRIPT LANGUAGE ="VBSCRIPT">
<!--
	Sub b1_Click()
		form1.submit		
	End Sub

	sub keypress()
			tBillno_keydown()
	end sub
-->
</Script>

<script language="JavaScript">
<!--
	function tBillno_keydown()
		{
		if (event.keyCode==13) 
			form1.submit() ;
	}
-->

</script>



<body background="image/blackground.gif">


<TABLE border=0 cellPadding=0 cellSpacing=0 height=55 width=836>
  <TR>
    <TD background=./image/top.gif colSpan=2 rowSpan=2 width=241>
      　
	</TD>
    <TD bgColor=#00b2eb height=1 width=591><IMG height=1 src="./image/pixel.gif" width=1></TD>
  </TR>
  <TR>
    <TD bgColor=#ffffff vAlign=bottom width="591">
      <DIV align=right style="width: 591; height: 64">
    <p align="center">
      <img border="0" src="image/buttons.gif" width="460" height="40"></p>
 
      <TABLE align=right border=0 cellPadding=0 cellSpacing=0 height=32 width=307 vspace="0" hspace="0">
        <TR>
          <TD background=./image/button1.gif width=161 align="left">
            <DIV align=center>
              <p align="center"><b><blink><font size="4">出 仓 管 理</font></blink></b></DIV>       
		  </TD>
          <TD vAlign=bottom width="142" align="left">
              <DIV align=right></DIV>
		  </TD>
		</TR>
	  </TABLE>
	  </DIV>
	</TD>
  </TR>
</TABLE>
<TABLE border=0 cellPadding=0 cellSpacing=0 height="317" width="1048">
  <TR>
    <TD width=249 height="1"></TD>
    <TD rowSpan=4 width=4 height="317">&nbsp;</TD>
    <TD height=1 width=789></TD>
  </TR>
  <TR>
    <TD vAlign=top width=249 height="318" rowspan="2">
      <TABLE align=right border=0 cellPadding=4 cellSpacing=4 height=270 width=140>
        <TR bgColor=#29a8cd>
          <TD width="122">
            <DIV align=center><b><a href="index.asp?op=reset"><font size="2" color="#FFFFFF">重新登录</font></a></b></DIV>
		  </TD>
		</TR>
        <TR bgColor=#2cb1bc>
          <TD width="122">
            <DIV align=center><b><a href="InStore.asp"><font color="#ffffff" size="2">进仓管理</font></a></b></DIV>
		  </TD>
		</TR>
        <TR bgColor=#35a8a5>
          <TD width="122">
            <DIV align=center><b><font color="#ffffff" size="2">出仓管理</font></b></DIV>
		  </TD>
		</TR>
        <TR bgColor=#35a88e>
          <TD width="122">
            <DIV align=center><b><a href="Client_List.htm"><font color="#ffffff" size="2">客户信息</font></a></b></DIV>
		  </TD>
		</TR>
        <TR bgColor=#37a67c>
          <TD width="122">
            <DIV align=center><b><a href="User.htm"><font color="#ffffff" size="2">用户设置</font></a></b></DIV>
		  </TD>
		</TR>
        <TR bgColor=#339966>
          <TD width="122">
            <DIV align=center><b><a href="jiliang_List.htm"><font color="#ffffff" size="2">选 
              项</font></a></b></DIV>
		  </TD>
		</TR>
        <TR bgColor=#339966>
          <TD width="122">
            <p align="center"><b><A href="report.htm"><font color="#ffffff" size="2">报表与计费</font></A></b>
		  </TD>
		</TR>

        <TR bgColor=#339966>
          <TD width="122">
            <p align="center"><b><A href="cremain.htm"><font color="#ffffff" size="2">结存情况</font></A></b>
		  </TD>
		</TR>

        <TR bgColor=#339966>
          <TD width="122">
            <p align="center"><b><A href="recover.htm"><font color="#ffffff" size="2">数据恢复</font></A></b>
		  </TD>
		</TR>
        <TR bgColor=#339966>
          <TD width="122">
            <DIV align=center><b><A href="support.htm"><font color="#ffffff" size="2">帮&nbsp;                                       
              助</font></A></b></DIV>
		  </TD>
		</TR>
	  </TABLE>
      <P align=right>　</P>
	</TD>
    <TD width="789" valign="top" height="45"> 
      <table border="0" width="532" height="29">
        <tr>
          <td width="168" background="image/button3.gif" height="25">
            <p align="center"><font face="宋体" size="3"><b><img border="0" src="image/stargold.gif" width="14" height="18">出仓选择</b></font></td>                                                                
          <td width="178" background="image/button3.gif" height="25">
            <p align="center"><font face="宋体" size="3"><b><a href="OutStore_Edit.asp">出仓编辑</a></b></font></td>                                                                
          <td width="178" background="image/button3.gif" height="25">
            <p align="right"><font face="宋体" size="3"><b><a href="Print_OutStore.asp">打印出库通知单</a></b></font></td>                                                                
        </tr>
      </table>
      </td>
  </tr>
  <TR>
    <TD valign="right" height="273" width="789"> 
    <% if session("varClientID") = "-" then
	 response.write	"<iframe src='OutStore_Browse_Frame.asp' height='600' width='"&iwidth&"' scrolling ='auto' frameborder='0'> </iframe>"
	   else 
	if session("varClientID") = "All" or session("varBillNo")<> "" or session("varContractID") <> "" or session("varSpecs") <> ""  or session("varGoodsName") <> "" then %>
  <% response.write  "<iframe src='OutStore_Browse_All_Frame.asp' height='600' width='"&iwidth&"' scrolling ='auto' frameborder='0'> </iframe>" %>     
    <% else 
    response.write	"<iframe src='OutStore_Browse_Frame.asp' height='600' width='"&iwidth&"' scrolling ='auto' frameborder='0'> </iframe>"
     end if
    end if	 %>
    
      </td>
  </tr>
  <tr> 
    <td width="249" height="30">　</td>
    <td valign="top" height="30" width="1000">     
      <form method="POST" action="OutStore.asp" name = "form1">
   
       <%
			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.open "select * from ClientInfo where ClientName='浙江中大技术进口有限公司' order by ClientName asc", conn, 1, 1	
		%>


       <p><font size="2"><b><font face="宋体">客户名称:</font>
       </b> 
        <select size="1" name="ClientName">                                     

<%        abcd="selected" %>
        <% do while not rs.EOF %>                                         
                <Option value = '<% = rs("ID") %>' <%if trim((Session("varClientID"))) = trim(CStr(rs("ID"))) then %> selected <% end if%> >
                <% =rs("ClientName") %>  </Option>
        <% rs.MoveNext 
           Loop %>
           <Option value ="All" <% if Session("varClientID") = "All" then %> selected <% end if %> >显示全部</Option> 
        <%   rs.Close %>
        </select>&nbsp; <b>进仓单号:</b><input type="text" name="tBillNo" size="8" onkeypress="keypress()"><b>&nbsp; 
       合同号:</b><input type="text" name="tContractID" size="7" onkeypress="keypress()"> <b> &nbsp;
	   品名:</b><input type="text" name="tGoodsName" size="4" onkeypress="keypress()"> <b> &nbsp;
	   规格:</b><input type="text" name="tSpecs" size="7" onkeypress="keypress()">
       </font>                                         
        <input type="button" value="浏  览" name="B1" OnClick ="b1_Click()"></p>                                       
      </form>
    </td>
  </tr>
</table>
</html>