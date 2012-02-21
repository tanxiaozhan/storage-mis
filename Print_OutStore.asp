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
	Session("varClientID")=Request.form("ClientName")
	Session("varBillNo")=Request.form("tBillNo")
	Session("varContractID")=Request.form("tContractID")
	
	if Request.form("R1")="V4" then
		Session("PrintChoice")="OutReceiveTable"
	else
		Session("PrintChoice")="OutStoreTable"
	end if
	
	
	
		''''''''''''''''''''''''''''''''''''''START修改毛重
		if Request.form("R2")="V5" then
		       
			        Session("PrintChoice_maozhong")="print_xianshimaozhong"
			   
			   
	    else
		  
			  Session("PrintChoice_maozhong")="print_buxianshimaozhong" 
			   
			   
	    end if
	
	''''''''''''''''''''''''''''''''''''''END
	
		
	
	
	Session("varStartYear")=Cint( request("StartYear") )
	Session("varStartMonth")=CInt( request("StartMonth") )
	Session("varStartDay")=CInt( request("StartDay") )
	Session("varEndYear")=CInt( request("EndYear") )
	Session("varEndMonth")=CInt( request("EndMonth") )
	Session("varEndDay")=CInt( request("EndDay") )
	
	
	if session("varStartYear")=0 then
		session("varStartYear") =  year(date()) 
	end if
	if session("varStartMonth")=0 then
		session("varStartMonth") = month(date())
	end if
	if session("varStartDay")=0 then
		session("varStartDay") = 1
	end if
	if session("varEndYear")=0 then
		session("varEndYear") = year(date())
	end if
	if session("varEndMonth")=0 then
		session("varEndMonth") = month(date())
	end if
	if session("varEndDay")=0 then
		session("varEndDay") = days( month( date() ) )
	end if
	
	
	
%>
 <SCRIPT LANGUAGE ="VBSCRIPT">
<!--
	Sub b1_Click()
		form1.submit		
	End Sub

	Sub OnR1Click()
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
<TABLE border=0 cellPadding=0 cellSpacing=0 height="305" width="1197">
  <TR>
    <TD width=249 height="1"></TD>
    <TD rowSpan=4 width=3 height="305">&nbsp;</TD>
    <TD height=1 width=935></TD>
  </TR>
  <TR>
    <TD vAlign=top width=249 height="308" rowspan="2">
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
      <P align=right>&nbsp;</P>
	</TD>
    <TD width="935" valign="top" height="35"> 
      <table border="0" width="533" height="29">
        <tr>
          <td width="173" background="image/button3.gif" height="25">
            <p align="center"><font face="宋体" size="3"><b><a href="OutStore.asp">出仓选择</a></b></font></td>                                                                
          <td width="176" background="image/button3.gif" height="25">
            <p align="center"><font face="宋体" size="3"><b><a href="OutStore_Edit.asp">出仓编辑</a></b></font></td>                                                                
          <td width="176" background="image/button3.gif" height="25">
            <p align="right"><font face="宋体" size="3"><b><img border="0" src="image/stargold.gif" width="14" height="18">打印出库通知单</b></font></td>                                                                
        </tr>
      </table>
      </td>
  </tr>
  <TR>
    <TD valign="right" height="273" width="935"> 
    <% if session("varClientID")="All" or session("varBillNo")<>"" or session("varContractID")<>"" then %>
    	<iframe src='OutStore_Print_All_Frame.asp' height='410' width='700'scrolling ="auto" frameborder="0"> </iframe>      
    
    <% else %>
    	<iframe src='OutStore_Print_Frame.asp' height='410' width='1000'scrolling ="auto" frameborder="0"> </iframe>      
    <% end if %>
      </td>
  </tr>
  <tr> 
    <td width="249" height="31">
      <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
    </td>
    <td valign="top" height="31" width="935">     
      <form method="POST" action="Print_OutStore.asp" name = "form1">
   
       <%
			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.open "select * from ClientInfo order by ClientName asc", conn, 1, 1	
		%>


       <p style="margin-top: 0; margin-bottom: 0"><b><font size="2"><font face="宋体">客户名称:</font> 
       </font></b><font size="2">   
        <select size="1" name="ClientName">                                            
        <Option value="-">----未选择----</Option>
<%        abcd="selected" %>
        <% do while not rs.EOF %>                                                  
                <Option value = '<% = rs("ID") %>' <%if trim((Session("varClientID"))) = trim(CStr(rs("ID"))) then %> selected <% end if%> >
                <% =rs("ClientName") %>  </Option>
        <% rs.MoveNext 
           Loop %>
           <Option value ="All" <% if Session("varClientID") = "All" then %> selected <% end if %> >显示全部</Option> 
        <%   rs.Close %>
        </select>&nbsp; <b>出仓单号:</b><input type="text" name="tBillNo" size="9" onkeypress="keypress()"><b>&nbsp; 
        合同号:</b><input type="text" name="tContractID" size="13" onkeypress="keypress()"></font>&nbsp;&nbsp;&nbsp;&nbsp;</p>                                    


       <p style="margin-top: 0; margin-bottom: 0"><font face="宋体" size="2"><b> <input type="radio"  value="V3" <%if Session("PrintChoice")="OutStoreTable" then %> checked <% end if %> name="R1" onclick="OnR1Click()"></b>出库通知单<b> 
       <input type="radio" name="R1" <%if Session("PrintChoice")="OutReceiveTable" then %> checked <% end if %>  value="V4" onclick="OnR1Click()"></b>货物签收单&nbsp;&nbsp; 
       <select size="1" name="StartYear">
        <%for i=year(date())-5 to year(date()) +1 %>
        		<option <% if session("varStartYear")=i then %> selected <% end if %> ><% = i %></option>
        <% next %>
       </select>年  
       <select size="1" name="StartMonth">
        <%for i=1 to 12 %>
        		<option <% if session("varStartMonth")=i then %> selected <% end if %> ><% = i %></option>
        <% next %>
       </select>月  
       <select size="1" name="StartDay">
        <%for i=1 to days( session("varStartMonth") ) %>
        		<option <% if session("varStartDay")=i then %> selected <% end if %> ><% = i %></option>
        <% next %>
       </select>日&nbsp; 至&nbsp; <select size="1" name="EndYear"> 
        <%for i=year(date())-5 to year(date()) +1 %>
        		<option <% if session("varEndYear")=i then %> selected <% end if %> ><% = i %></option>
        <% next %>
       </select>年  
       <select size="1" name="EndMonth">
        <%for i=1 to 12 %>
        		<option <% if session("varEndMonth")=i then %> selected <% end if %> ><% = i %></option>
        <% next %>
       </select>月  
       <select size="1" name="EndDay">
        <%for i=1 to days( session("varEndMonth") ) %>
        		<option <% if session("varEndDay")=i then %> selected <% end if %> ><% = i %></option>
        <% next %>
       </select>日 </font><input type="button" value="浏  览" name="B1" OnClick ="b1_Click()"></p>   
	   

	    <input type="radio"  value="V6" <%if Session("PrintChoice_maozhong")="print_buxianshimaozhong" then %> checked <% end if %> name="R2" onclick="OnR1Click()"></b><font color="#ff2323">不显示毛重</font><b> 
        <input type="radio" name="R2" <%if Session("PrintChoice_maozhong")="print_xianshimaozhong" then %> checked <% end if %>  value="V5" onclick="OnR1Click()"></b><font color="#ff2323">显示毛重</font>&nbsp;&nbsp;
  
	   
	   

	   
      </form>
    </td>
  </tr>
</table>
</html>
